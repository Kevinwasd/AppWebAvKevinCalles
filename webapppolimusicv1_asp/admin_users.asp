<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/sha1.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/upload.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<%
' =====================================================================
'  PoliMusic - Administracion de usuarios (CRUD completo)
'  Roles segun TBL_CATALOG (COD_CATALOG_PARENT = '1'):
'     0 = USER (usuario normal)   1 = ADMIN (administrador)
' =====================================================================
Dim sAction, nId, sMsg, bOk, sUser, sEmail, sPass, sPass2, sBirth, sPhoto
Dim nType, nNewId, sSearch, nPage, nTotal, nPages, rs, bForm, nEdit
Dim sFUser, sFEmail, sFBirth, sFPhoto, nFType, sWhere, nOthers

RequireAdmin()

' --------------------------- Acciones POST ---------------------------
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sAction = LCase(UFField("action", ""))
    If UFField("do_delete", "") = "1" Then sAction = "delete"
    nId = UFFieldInt("id", 0)

    If sAction = "delete" Then
        If nId <= 0 Then
            SetFlash "Usuario no valido.", "error"
        ElseIf nId = CurrentUserId() Then
            SetFlash "No puede eliminar su propia cuenta mientras la usa.", "error"
        Else
            ' No permitir quedarse sin administradores
            nOthers = CLng(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE USER_TYPE = " & ROLE_ADMIN _
                & " AND ID_USER <> " & nId, 0))
            If CLng(DbScalar("SELECT ISNULL(MIN(USER_TYPE),0) FROM TBL_USER WHERE ID_USER = " & nId, 0)) = ROLE_ADMIN _
               And nOthers = 0 Then
                SetFlash "No es posible eliminar al ultimo administrador del sistema.", "error"
            Else
                sPhoto = DbScalar("SELECT ISNULL(MIN(USER_PHOTO),'') FROM TBL_USER WHERE ID_USER = " & nId, "")
                DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN " _
                    & "(SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = " & nId & ")"
                DbExec "DELETE FROM TBL_PLAYLIST WHERE ID_USER = " & nId
                DbExec "DELETE FROM TBL_USER WHERE ID_USER = " & nId
                If Len(sPhoto & "") > 0 Then DeleteAsset sPhoto
                SetFlash "El usuario fue eliminado correctamente.", "success"
            End If
        End If
        Response.Redirect U("admin_users.asp")
    End If

    If sAction = "save" Then
        sUser  = Left(UFField("username", ""), 50)
        sEmail = Left(UFField("email", ""), 50)
        sPass  = UFField("password", "")
        sPass2 = UFField("password2", "")
        sBirth = UFField("birthday", "")
        nType  = UFFieldInt("user_type", ROLE_USER)
        bOk = True
        sMsg = ""

        If nType <> ROLE_ADMIN Then nType = ROLE_USER

        If Len(sUser) < 4 Then
            sMsg = "El nombre de usuario debe tener al menos 4 caracteres."
            bOk = False
        ElseIf Len(sEmail) < 6 Or InStr(sEmail, "@") = 0 Then
            sMsg = "Debe indicar un correo electronico valido."
            bOk = False
        ElseIf nId <= 0 And Len(sPass) < 6 Then
            sMsg = "La contrasena es obligatoria para un usuario nuevo (minimo 6 caracteres)."
            bOk = False
        ElseIf Len(sPass) > 0 And Len(sPass) < 6 Then
            sMsg = "La contrasena debe tener al menos 6 caracteres."
            bOk = False
        ElseIf Len(sPass) > 0 And sPass <> sPass2 Then
            sMsg = "Las contrasenas no coinciden."
            bOk = False
        ElseIf Len(sBirth) > 0 And Not IsDate(sBirth) Then
            sMsg = "La fecha de nacimiento no es valida."
            bOk = False
        ElseIf UsernameExists(sUser, nId) Then
            sMsg = "El nombre de usuario ya esta registrado."
            bOk = False
        ElseIf EmailExists(sEmail, nId) Then
            sMsg = "El correo electronico ya esta registrado."
            bOk = False
        ElseIf nId = CurrentUserId() And nType <> ROLE_ADMIN Then
            sMsg = "No puede quitarse a si mismo el rol de administrador."
            bOk = False
        End If

        If bOk Then
            If nId > 0 Then
                If CLng(DbScalar("SELECT COUNT(*) FROM TBL_USER WHERE ID_USER = " & nId, 0)) = 0 Then
                    bOk = False
                    sMsg = "El usuario indicado no existe."
                Else
                    DbExec "UPDATE TBL_USER SET USERNAME = " & SqlStr(sUser) _
                        & ", EMAIL = " & SqlStr(sEmail) _
                        & ", BIRTHDAY = " & SqlDate(sBirth) _
                        & ", USER_TYPE = " & nType & " WHERE ID_USER = " & nId
                End If
            Else
                nNewId = DbInsert("INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, " _
                    & "USER_PHOTO, USER_TYPE) VALUES (" & SqlStr(sUser) & ", " _
                    & SqlStr(Sha1Hex(sPass)) & ", " & SqlStr(sEmail) & ", " & SqlDate(sBirth) & ", " _
                    & SqlStr("~/UserImages/default.jpg") & ", " & nType & ")")
                If nNewId <= 0 Then
                    bOk = False
                    sMsg = "No fue posible crear el usuario."
                Else
                    nId = nNewId
                End If
            End If
        End If

        If bOk And Len(sPass) > 0 And nId > 0 Then
            DbExec "UPDATE TBL_USER SET PASSWORD = " & SqlStr(Sha1Hex(sPass)) & " WHERE ID_USER = " & nId
        End If

        If bOk And HasUpload("photo") Then
            sPhoto = SaveUploaded("photo", DIR_USERS, "user_" & nId)
            If Len(sPhoto) > 0 Then
                DbExec "UPDATE TBL_USER SET USER_PHOTO = " & SqlStr(sPhoto) & " WHERE ID_USER = " & nId
            Else
                bOk = False
                sMsg = "No fue posible guardar la foto del usuario."
            End If
        End If

        If bOk Then
            If nId = CurrentUserId() Then RefreshSession
            SetFlash "El usuario fue guardado correctamente.", "success"
            Response.Redirect U("admin_users.asp")
        Else
            SetFlash sMsg, "error"
        End If
    End If
End If

nEdit   = ParamInt("edit", 0)
bForm   = (nEdit > 0) Or (Param("new", "") = "1")
sSearch = ParamMax("q", 60)
nPage   = ParamInt("page", 1)

PAGE_TITLE   = "Administrar usuarios"
ACTIVE_ADMIN = "users"
%>
<!--#include file="includes/layout_top.asp"-->

<% If bForm Then
    sFUser = ""
    sFEmail = ""
    sFBirth = ""
    sFPhoto = ""
    nFType = ROLE_USER
    If nEdit > 0 Then
        Set rs = DbQuery("SELECT TOP 1 USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE " _
            & "FROM TBL_USER WHERE ID_USER = " & nEdit)
        If Not rs.EOF Then
            sFUser  = rs("USERNAME") & ""
            sFEmail = rs("EMAIL") & ""
            sFBirth = FmtDateInput(rs("BIRTHDAY"))
            sFPhoto = rs("USER_PHOTO") & ""
            nFType  = CLng(NumOr(rs("USER_TYPE"), ROLE_USER))
        End If
        rs.Close
        Set rs = Nothing
    End If
%>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>
                <% If nEdit > 0 Then %>Editar usuario<% Else %>Nuevo usuario<% End If %></h2>
            <p class="section-sub">Los roles se administran desde la tabla TBL_CATALOG</p>
        </div>
        <a class="btn btn-ghost btn-sm" href="<%=U("admin_users.asp")%>">
            <%=Icon("chevron-left")%><span>Volver al listado</span></a>
        </div>

        <form method="post" action="<%=U("admin_users.asp")%>" enctype="multipart/form-data">
            <input type="hidden" name="action" value="save">
            <input type="hidden" name="id" value="<%=nEdit%>">

            <div class="panel">
                <h3 class="panel-title"><%=Icon("user")%> Datos de la cuenta</h3>
                <div class="form-grid">
                    <div class="field">
                        <label for="username">Nombre de usuario *</label>
                        <input class="input" type="text" id="username" name="username"
                               maxlength="50" required value="<%=H(sFUser)%>" placeholder="Ej. epn.fis">
                    </div>
                    <div class="field">
                        <label for="email">Correo electronico *</label>
                        <input class="input" type="email" id="email" name="email"
                               maxlength="50" required value="<%=H(sFEmail)%>">
                    </div>
                    <div class="field">
                        <label for="user_type">Rol / tipo de usuario</label>
                        <select class="select" id="user_type" name="user_type">
                            <%
                            Set rs = DbQuery("SELECT COD_CATALOG, [VALUE] FROM TBL_CATALOG " _
                                & "WHERE COD_CATALOG_PARENT = '1' ORDER BY COD_CATALOG")
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
                        <label for="birthday">Fecha de nacimiento</label>
                        <input class="input" type="date" id="birthday" name="birthday"
                               value="<%=H(sFBirth)%>">
                    </div>
                </div>
            </div>

            <div class="panel">
                <h3 class="panel-title"><%=Icon("lock")%> Credenciales y foto</h3>
                <div class="form-grid">
                    <div class="field">
                        <label for="password">
                            Contrasena <% If nEdit > 0 Then %><span class="hint">(vacio = sin cambios)</span>
                            <% Else %><span class="hint">*</span><% End If %>
                        </label>
                        <input class="input" type="password" id="password" name="password"
                               maxlength="50" autocomplete="new-password">
                    </div>
                    <div class="field">
                        <label for="password2">Repetir contrasena</label>
                        <input class="input" type="password" id="password2" name="password2"
                               maxlength="50" autocomplete="new-password">
                    </div>
                    <div class="field">
                        <label for="photo">Foto de perfil <span class="hint">(imagen)</span></label>
                        <input class="input" type="file" id="photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                        <span class="hint">Ruta actual: <%=H(IIf2(Len(sFPhoto) > 0, sFPhoto, "sin imagen"))%></span>
                    </div>
                    <div class="field">
                        <label>Vista previa</label>
                        <img style="width:76px;height:76px;border-radius:50%;object-fit:cover;border:1px solid var(--pm-border-2)"
                             alt="" src="<%=H(ImgUrl(sFPhoto, "user", IIf2(nEdit > 0, nEdit, 99), FirstLetter(sFUser)))%>">
                    </div>
                </div>
                <p class="small muted mb-0">
                    Las contrasenas se almacenan con hash SHA-1 (40 caracteres hexadecimales),
                    por lo que nunca es posible recuperar la original.
                </p>
            </div>

            <div class="form-actions">
                <button type="submit" class="btn btn-primary">
                    <%=Icon("check")%><span>Guardar usuario</span></button>
                <a class="btn btn-ghost" href="<%=U("admin_users.asp")%>">Cancelar</a>
                <% If nEdit > 0 And nEdit <> CurrentUserId() Then %>
                <button type="submit" name="do_delete" value="1" class="btn btn-danger"
                        data-confirm="&iquest;Eliminar este usuario y todas sus playlists?">
                    <%=Icon("trash")%><span>Eliminar</span></button>
                <% End If %>
            </div>
        </form>
    </section>
<% Else %>
    <section class="section">
        <div class="section-head"><div>
            <h2 class="section-title"><span class="dot"></span>Usuarios</h2>
            <p class="section-sub">Administracion de cuentas y roles</p>
        </div>
        <a class="btn btn-primary btn-sm" href="<%=U("admin_users.asp?new=1")%>">
            <%=Icon("plus")%><span>Nuevo usuario</span></a>
        </div>

        <form method="get" action="<%=U("admin_users.asp")%>" class="toolbar">
            <input class="filter" type="search" name="q" value="<%=H(sSearch)%>"
                   placeholder="Buscar por usuario o correo..." autocomplete="off">
            <button type="submit" class="btn btn-outline btn-sm"><%=Icon("search")%><span>Buscar</span></button>
            <% If Len(sSearch) > 0 Then %>
            <a class="btn btn-ghost btn-sm" href="<%=U("admin_users.asp")%>">Limpiar</a>
            <% End If %>
        </form>

        <%
        sWhere = ""
        If Len(sSearch) > 0 Then
            sWhere = "WHERE (USERNAME LIKE " & SqlStr("%" & sSearch & "%") _
                & " OR EMAIL LIKE " & SqlStr("%" & sSearch & "%") & ")"
        End If
        nTotal = CLng(DbScalar("SELECT COUNT(*) FROM TBL_USER " & sWhere, 0))
        nPages = PageCount(nTotal, PAGE_SIZE)
        If nPages < 1 Then nPages = 1
        If nPage < 1 Then nPage = 1
        If nPage > nPages Then nPage = nPages

        Set rs = DbQuery("SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, " _
            & "(SELECT COUNT(*) FROM TBL_PLAYLIST pl WHERE pl.ID_USER = u.ID_USER) AS N_PL " _
            & "FROM TBL_USER u " & sWhere & " ORDER BY u.ID_USER ASC " _
            & "OFFSET " & PageOffset(nPage, PAGE_SIZE) & " ROWS FETCH NEXT " & PAGE_SIZE & " ROWS ONLY")
        %>
        <div class="table-wrap">
            <table class="table">
                <thead><tr>
                    <th style="width:64px">ID</th>
                    <th>Usuario</th>
                    <th class="hide-sm">Correo</th>
                    <th>Rol</th>
                    <th class="hide-sm">Nacimiento</th>
                    <th>Playlists</th>
                    <th class="actions">Acciones</th>
                </tr></thead>
                <tbody>
                <% If rs.EOF Then %>
                    <tr><td colspan="7" class="muted">No se encontraron usuarios.</td></tr>
                <% End If %>
                <%
                Do While Not rs.EOF
                %>
                    <tr>
                        <td class="muted">#<%=CLng(rs("ID_USER"))%></td>
                        <td>
                            <div class="cell-flex">
                                <img class="table-avatar" alt=""
                                     src="<%=H(ImgUrl(rs("USER_PHOTO"), "user", rs("ID_USER"), FirstLetter(rs("USERNAME"))))%>">
                                <span class="cell-strong"><%=H(rs("USERNAME"))%>
                                    <% If CLng(rs("ID_USER")) = CurrentUserId() Then %>
                                    <span class="badge gray">usted</span>
                                    <% End If %>
                                </span>
                            </div>
                        </td>
                        <td class="hide-sm muted small"><%=H(rs("EMAIL"))%></td>
                        <td>
                            <% If CLng(NumOr(rs("USER_TYPE"), 0)) = ROLE_ADMIN Then %>
                            <span class="badge">Administrador</span>
                            <% Else %>
                            <span class="badge gray">Usuario</span>
                            <% End If %>
                        </td>
                        <td class="hide-sm muted small"><%=FmtDate(rs("BIRTHDAY"))%></td>
                        <td><span class="badge gray"><%=FmtNum(rs("N_PL"))%></span></td>
                        <td class="actions">
                            <a class="btn-icon" title="Editar"
                               href="<%=U("admin_users.asp?edit=" & rs("ID_USER"))%>"><%=Icon("edit")%></a>
                            <% If CLng(rs("ID_USER")) <> CurrentUserId() Then %>
                            <form method="post" style="display:inline" action="<%=U("admin_users.asp")%>">
                                <input type="hidden" name="action" value="delete">
                                <input type="hidden" name="id" value="<%=CLng(rs("ID_USER"))%>">
                                <button type="submit" class="btn-icon danger" title="Eliminar"
                                        data-confirm="&iquest;Eliminar este usuario y todas sus playlists?"><%=Icon("trash")%></button>
                            </form>
                            <% End If %>
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
        Pager nPage, nPages, U("admin_users.asp" & IIf2(Len(sSearch) > 0, "?q=" & UrlEnc(sSearch), ""))
        %>
    </section>
<% End If %>

<!--#include file="includes/layout_bottom.asp"-->
