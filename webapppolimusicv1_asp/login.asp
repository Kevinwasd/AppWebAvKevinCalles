<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/sha1.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/icons.asp"-->
<%
' =====================================================================
'  PoliMusic - Inicio de sesion
' =====================================================================
Dim sUser, sPass, sErr, sNext, rs, sOk
If IsLoggedIn() Then
    Response.Redirect U("home.asp")
End If
sErr  = ""
sOk   = Param("ok", "")
sNext = SafeLocalUrl(Param("next", ""))
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    sUser = ParamMax("username", 50)
    sPass = ParamMax("password", 50)
    If Len(sUser) = 0 Or Len(sPass) = 0 Then
        sErr = "Debe indicar el usuario y la contrasena."
    Else
        Set rs = UserByLogin(sUser)
        If rs.EOF Then
            sErr = "El usuario indicado no existe."
        ElseIf Not PasswordMatches(sPass, rs("PASSWORD") & "") Then
            sErr = "La contrasena no es correcta."
        Else
            LoginUser rs
            rs.Close
            Set rs = Nothing
            If Len(sNext) = 0 Then sNext = "home.asp"
            Response.Redirect U(sNext)
        End If
        If Not rs Is Nothing Then
            If rs.State <> 0 Then rs.Close
            Set rs = Nothing
        End If
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Iniciar sesion &middot; <%=H(APP_NAME)%></title>
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
        <h2>Toda tu musica,<br>en un solo lugar.</h2>
        <p><%=H(APP_TAGLINE)%>. Escuche canciones, organice sus playlists y descubra
           artistas y albumes desde cualquier navegador.</p>
        <ul class="auth-features">
            <li><span class="tick">&check;</span> Reproductor con controles completos y cola de reproduccion</li>
            <li><span class="tick">&check;</span> Playlists publicas y privadas, totalmente administrables</li>
            <li><span class="tick">&check;</span> Catalogo de artistas y albumes con portadas</li>
            <li><span class="tick">&check;</span> Panel de administracion con operaciones CRUD</li>
        </ul>
    </section>
    <section class="auth-panel">
        <div class="auth-card">
            <h1>Iniciar sesion</h1>
            <p class="lead">Ingrese sus credenciales para continuar.</p>
            <% If Len(sOk) > 0 Then %>
            <div class="flash flash-success">
                <span class="flash-icon"></span>
                <span>Su cuenta fue creada correctamente. Ya puede iniciar sesion.</span>
            </div>
            <% End If %>
            <% If Len(sErr) > 0 Then %>
            <div class="flash flash-error">
                <span class="flash-icon"></span>
                <span><%=H(sErr)%></span>
            </div>
            <% End If %>
            <form method="post" action="<%=U("login.asp")%>" autocomplete="off">
                <input type="hidden" name="next" value="<%=H(sNext)%>">
                <div class="field">
                    <label for="username">Usuario</label>
                    <input class="input" type="text" id="username" name="username"
                           value="<%=H(sUser)%>" maxlength="50" required autofocus
                           placeholder="Ej. userAdmin">
                </div>
                <div class="field">
                    <label for="password">Contrasena</label>
                    <input class="input" type="password" id="password" name="password"
                           maxlength="50" required placeholder="Su contrasena">
                </div>
                <button type="submit" class="btn btn-primary btn-block mt-16">
                    <%=Icon("logout")%> <span>Entrar a <%=H(APP_NAME)%></span>
                </button>
            </form>
            <div class="auth-foot">
                &iquest;No tiene una cuenta? <a href="<%=U("register.asp")%>">Registrese aqui</a>
            </div>
            <div class="demo-box mt-24">
                <strong>Cuentas de prueba</strong>
                <div class="demo-row"><span>Administrador</span>
                    <span><code>userAdmin</code> / <code>12345678</code></span></div>
                <div class="demo-row"><span>Usuario normal</span>
                    <span><code>epn.fis</code> / <code>12345678</code></span></div>
                <div class="demo-row"><span>Usuario normal</span>
                    <span><code>userNormal</code> / <code>12345678</code></span></div>
            </div>
        </div>
    </section>
</div>
</body>
</html>
