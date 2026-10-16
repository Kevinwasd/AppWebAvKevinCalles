<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/sha1.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/upload.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<!--#include file="includes/playlists_logic.asp"-->
<!--#include file="includes/actions.asp"-->
<%
' =====================================================================
'  PoliMusic - Mi perfil
' =====================================================================
Dim rs, sUser, sEmail, sBirth, nPl, nPlSongs, nPlays

RequireLogin()
HandlePostActions()

PAGE_TITLE = "Mi perfil"
VIEW_NAME  = "profile"

Set rs = UserById(CurrentUserId())
sUser  = rs("USERNAME") & ""
sEmail = rs("EMAIL") & ""
sBirth = FmtDateInput(rs("BIRTHDAY"))
rs.Close
Set rs = Nothing

nPl      = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = " & CurrentUserId(), 0))
nPlSongs = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps INNER JOIN TBL_PLAYLIST p " _
    & "ON p.ID_PLAYLIST = ps.ID_PLAYLIST WHERE p.ID_USER = " & CurrentUserId(), 0))
nPlays   = CLng(DbScalar("SELECT ISNULL(SUM(PLAYS),0) FROM TBL_SONG", 0))
%>
<!--#include file="includes/layout_top.asp"-->

    <div class="detail-head">
        <img class="detail-cover round" alt="<%=H(sUser)%>"
             src="<%=H(ImgUrl(CurrentUserPhoto(), "user", CurrentUserId(), FirstLetter(sUser)))%>">
        <div class="detail-info">
            <span class="kind"><% If IsAdmin() Then %>Administrador<% Else %>Usuario<% End If %></span>
            <h1><%=H(sUser)%></h1>
            <div class="detail-meta">
                <span><%=H(sEmail)%></span>
                <span class="sep">&middot;</span>
                <span>Cuenta #<%=CurrentUserId()%></span>
            </div>
        </div>
    </div>

    <div class="stats">
        <div class="stat">
            <span class="ic"><%=Icon("playlist")%></span>
            <span><b><%=FmtNum(nPl)%></b><span>Playlists propias</span></span>
        </div>
        <div class="stat">
            <span class="ic"><%=Icon("music")%></span>
            <span><b><%=FmtNum(nPlSongs)%></b><span>Canciones en sus playlists</span></span>
        </div>
        <div class="stat">
            <span class="ic"><%=Icon("chart")%></span>
            <span><b><%=FmtNum(nPlays)%></b><span>Reproducciones del sitio</span></span>
        </div>
    </div>

    <section class="section">
        <div class="panel">
            <h3 class="panel-title"><%=Icon("edit")%> Datos de la cuenta</h3>
            <form method="post" action="<%=U("profile.asp")%>" enctype="multipart/form-data">
                <input type="hidden" name="action" value="profile_save">
                <div class="form-grid">
                    <div class="field">
                        <label for="p_user">Nombre de usuario</label>
                        <input class="input" type="text" id="p_user" name="username"
                               maxlength="50" required value="<%=H(sUser)%>">
                    </div>
                    <div class="field">
                        <label for="p_email">Correo electronico</label>
                        <input class="input" type="email" id="p_email" name="email"
                               maxlength="50" required value="<%=H(sEmail)%>">
                    </div>
                    <div class="field">
                        <label for="p_birth">Fecha de nacimiento</label>
                        <input class="input" type="date" id="p_birth" name="birthday"
                               value="<%=H(sBirth)%>">
                    </div>
                    <div class="field">
                        <label for="p_photo">Foto de perfil <span class="hint">(deje vacio para conservarla)</span></label>
                        <input class="input" type="file" id="p_photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                    </div>
                </div>

                <h3 class="panel-title mt-24"><%=Icon("lock")%> Cambiar contrasena</h3>
                <p class="small muted mt-0 mb-16">
                    Deje estos campos vacios si no desea modificar su contrasena.
                    Las contrasenas se almacenan con hash SHA-1.
                </p>
                <div class="form-grid">
                    <div class="field">
                        <label for="p_pass">Nueva contrasena</label>
                        <input class="input" type="password" id="p_pass" name="password"
                               minlength="6" maxlength="50" autocomplete="new-password">
                    </div>
                    <div class="field">
                        <label for="p_pass2">Repetir nueva contrasena</label>
                        <input class="input" type="password" id="p_pass2" name="password2"
                               minlength="6" maxlength="50" autocomplete="new-password">
                    </div>
                </div>

                <div class="form-actions">
                    <button type="submit" class="btn btn-primary">
                        <%=Icon("check")%><span>Guardar cambios</span></button>
                    <a class="btn btn-ghost" href="<%=U("home.asp")%>">Cancelar</a>
                </div>
            </form>
        </div>
    </section>

    <section class="section">
        <div class="section-head">
            <div>
                <h2 class="section-title"><span class="dot"></span>Mis playlists</h2>
                <p class="section-sub">Listas creadas con esta cuenta</p>
            </div>
            <a class="link-more" href="<%=U("home.asp?view=playlists&mine=1")%>" data-nav>Administrar &rarr;</a>
        </div>
        <%
        Set rs = DbQuery("SELECT TOP 6 pl.ID_PLAYLIST, pl.TITLE, pl.PHOTO, pl.TYPE, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = pl.ID_PLAYLIST) AS N_SONGS " _
            & "FROM TBL_PLAYLIST pl WHERE pl.ID_USER = " & CurrentUserId() & " ORDER BY pl.ID_PLAYLIST DESC")
        If rs.EOF Then
            EmptyState "Sin playlists", "Todavia no ha creado ninguna playlist."
        Else
        %>
        <div class="grid">
        <%
            Do While Not rs.EOF
                PlaylistCard rs("ID_PLAYLIST"), rs("TITLE") & "", rs("PHOTO") & "", _
                             rs("TYPE"), rs("N_SONGS"), sUser
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

<!--#include file="includes/layout_bottom.asp"-->
