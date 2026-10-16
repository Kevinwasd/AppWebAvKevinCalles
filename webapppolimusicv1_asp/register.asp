<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/sha1.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/upload.asp"-->
<!--#include file="includes/icons.asp"-->
<%
' =====================================================================
'  PoliMusic - Registro de nuevos usuarios
'  Todo usuario registrado desde aqui obtiene el rol de USUARIO (0).
' =====================================================================
Dim sUser, sEmail, sPass, sPass2, sBirth, sErr, sPhoto, nNewId
Dim aErr
If IsLoggedIn() Then
    Response.Redirect U("home.asp")
End If
sErr   = ""
sUser  = ""
sEmail = ""
sBirth = ""
aErr   = 0
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sUser  = Left(UFField("username", ""), 50)
    sEmail = Left(UFField("email", ""), 50)
    sPass  = UFField("password", "")
    sPass2 = UFField("password2", "")
    sBirth = UFField("birthday", "")
    If Len(sUser) < 4 Then
        sErr = "El nombre de usuario debe tener al menos 4 caracteres."
    ElseIf Len(sEmail) < 6 Or InStr(sEmail, "@") = 0 Or InStr(sEmail, ".") = 0 Then
        sErr = "Debe indicar un correo electronico valido."
    ElseIf Len(sPass) < 6 Then
        sErr = "La contrasena debe tener al menos 6 caracteres."
    ElseIf sPass <> sPass2 Then
        sErr = "Las contrasenas no coinciden."
    ElseIf UsernameExists(sUser, 0) Then
        sErr = "El nombre de usuario ya esta registrado."
    ElseIf EmailExists(sEmail, 0) Then
        sErr = "El correo electronico ya esta registrado."
    ElseIf Len(sBirth) > 0 And Not IsDate(sBirth) Then
        sErr = "La fecha de nacimiento no es valida."
    ElseIf Len(sBirth) > 0 And CDate(sBirth) > Date() Then
        sErr = "La fecha de nacimiento no puede ser futura."
    Else
        sPhoto = "~/UserImages/default.jpg"
        nNewId = DbInsert("INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, " _
            & "USER_PHOTO, USER_TYPE) VALUES (" _
            & SqlStr(sUser) & ", " & SqlStr(Sha1Hex(sPass)) & ", " & SqlStr(sEmail) & ", " _
            & SqlDate(sBirth) & ", " & SqlStr(sPhoto) & ", " & ROLE_USER & ")")
        If nNewId <= 0 Then
            sErr = "No fue posible crear la cuenta. Intente nuevamente."
        Else
            ' Si el usuario adjunto una foto, se guarda y se asocia a la cuenta.
            If HasUpload("photo") Then
                sPhoto = SaveUploaded("photo", DIR_USERS, "user_" & nNewId)
                If Len(sPhoto) > 0 Then
                    DbExec "UPDATE TBL_USER SET USER_PHOTO = " & SqlStr(sPhoto) _
                        & " WHERE ID_USER = " & nNewId
                End If
            End If
            Response.Redirect U("login.asp?ok=1")
        End If
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Crear cuenta &middot; <%=H(APP_NAME)%></title>
<link rel="icon" href="<%=U("assets/img/favicon.svg")%>" type="image/svg+xml">
<link rel="stylesheet" href="<%=U("assets/css/style.css")%>?v=<%=H(APP_VERSION)%>">
</head>
<body>
<div class="auth-wrap">
    <section class="auth-side">
        <a class="brand" href="<%=U("login.asp")%>">
            <span class="brand-logo"><%=LogoMark()%></span>
            <span class="brand-text">
                <strong><%=H(APP_NAME)%></strong>
                <span>Musica en linea</span>
            </span>
        </a>
        <h2>Cree su cuenta<br>y empiece a escuchar.</h2>
        <p>Registrarse es inmediato. Con su cuenta podra crear playlists
           publicas o privadas y guardar sus canciones favoritas.</p>
        <ul class="auth-features">
            <li><span class="tick">&check;</span> Playlists ilimitadas</li>
            <li><span class="tick">&check;</span> Contrasena protegida con SHA-1</li>
            <li><span class="tick">&check;</span> Foto de perfil personalizable</li>
        </ul>
    </section>
    <section class="auth-panel">
        <div class="auth-card">
            <h1>Crear cuenta</h1>
            <p class="lead">Complete el formulario para registrarse como usuario.</p>
            <% If Len(sErr) > 0 Then %>
            <div class="flash flash-error">
                <span class="flash-icon"></span>
                <span><%=H(sErr)%></span>
            </div>
            <% End If %>
            <form method="post" action="<%=U("register.asp")%>"
                  enctype="multipart/form-data" autocomplete="off">
                <div class="field">
                    <label for="username">Nombre de usuario <span class="hint">(minimo 4 caracteres)</span></label>
                    <input class="input" type="text" id="username" name="username"
                           value="<%=H(sUser)%>" maxlength="50" required autofocus
                           placeholder="Ej. epn.fis">
                </div>
                <div class="field">
                    <label for="email">Correo electronico</label>
                    <input class="input" type="email" id="email" name="email"
                           value="<%=H(sEmail)%>" maxlength="50" required
                           placeholder="usuario@epn.edu.ec">
                </div>
                <div class="form-grid">
                    <div class="field">
                        <label for="password">Contrasena</label>
                        <input class="input" type="password" id="password" name="password"
                               minlength="6" maxlength="50" required>
                    </div>
                    <div class="field">
                        <label for="password2">Repetir contrasena</label>
                        <input class="input" type="password" id="password2" name="password2"
                               minlength="6" maxlength="50" required>
                    </div>
                </div>
                <div class="form-grid">
                    <div class="field">
                        <label for="birthday">Fecha de nacimiento <span class="hint">(opcional)</span></label>
                        <input class="input" type="date" id="birthday" name="birthday"
                               value="<%=H(sBirth)%>">
                    </div>
                    <div class="field">
                        <label for="photo">Foto de perfil <span class="hint">(opcional, imagen)</span></label>
                        <input class="input" type="file" id="photo" name="photo"
                               accept="image/png,image/jpeg,image/gif,image/webp">
                    </div>
                </div>
                <button type="submit" class="btn btn-primary btn-block mt-16">
                    <%=Icon("check")%> <span>Crear mi cuenta</span>
                </button>
            </form>
            <div class="auth-foot">
                &iquest;Ya tiene una cuenta? <a href="<%=U("login.asp")%>">Inicie sesion</a>
            </div>
        </div>
    </section>
</div>
</body>
</html>
