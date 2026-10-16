<%
' =====================================================================
'  PoliMusic - Cabecera de pagina (shell): barra lateral + barra superior
'  Requiere que la pagina defina PAGE_TITLE y, opcionalmente, VIEW_NAME
'  y ACTIVE_ADMIN antes de incluir este archivo.
' =====================================================================
Dim PAGE_TITLE_V, VIEW_NAME_V, ACTIVE_ADMIN_V
If Len(PAGE_TITLE & "") = 0 Then
    PAGE_TITLE_V = APP_NAME
Else
    PAGE_TITLE_V = PAGE_TITLE
End If
VIEW_NAME_V = VIEW_NAME & ""
ACTIVE_ADMIN_V = ACTIVE_ADMIN & ""

' Elemento de navegacion principal.
Sub NavMain(ByVal sView, ByVal sLabel, ByVal sIcon)
    Dim cls
    cls = "nav-item"
    If LCase(VIEW_NAME_V) = LCase(sView) Then cls = cls & " active"
    Response.Write "<a class=""" & cls & """ href=""" & U("home.asp?view=" & sView) & """"_
        & " data-nav=""" & H(sView) & """>" & Icon(sIcon) & "<span>" & H(sLabel) & "</span></a>"
End Sub

' Elemento de navegacion de administracion.
Sub NavAdmin(ByVal sKey, ByVal sFile, ByVal sLabel, ByVal sIcon)
    Dim cls
    cls = "nav-item"
    If LCase(ACTIVE_ADMIN_V) = LCase(sKey) Then cls = cls & " active"
    Response.Write "<a class=""" & cls & """ href=""" & U(sFile) & """>"_
        & Icon(sIcon) & "<span>" & H(sLabel) & "</span></a>"
End Sub
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="PoliMusic - plataforma de musica en linea">
<title><%=H(PAGE_TITLE_V)%> &middot; <%=H(APP_NAME)%></title>
<link rel="icon" href="<%=U("assets/img/favicon.svg")%>" type="image/svg+xml">
<link rel="stylesheet" href="<%=U("assets/css/style.css")%>?v=<%=H(APP_VERSION)%>">
</head>
<body>
<div class="app">

    <!-- ============================ BARRA LATERAL ==================== -->
    <aside class="sidebar" id="pmSidebar">
        <a class="brand" href="<%=U("home.asp")%>" data-nav>
            <span class="brand-logo"><%=LogoMark()%></span>
            <span class="brand-text">
                <strong><%=H(APP_NAME)%></strong>
                <span>Musica en linea</span>
            </span>
        </a>

        <nav class="nav">
            <div class="nav-title">Explorar</div>
            <% NavMain "home", "Inicio", "home" %>
            <% NavMain "songs", "Canciones", "music" %>
            <% NavMain "albums", "Albumes", "album" %>
            <% NavMain "artists", "Artistas", "artist" %>
            <% NavMain "playlists", "Playlists", "playlist" %>
        </nav>

        <nav class="nav">
            <div class="nav-title">Mi musica</div>
            <a class="nav-item" href="<%=U("home.asp?view=playlists&mine=1")%>" data-nav="playlists">
                <%=Icon("user")%><span>Mis playlists</span>
            </a>
            <a class="nav-item" href="<%=U("profile.asp")%>">
                <%=Icon("edit")%><span>Mi perfil</span>
            </a>
        </nav>

        <% If IsAdmin() Then %>
        <nav class="nav">
            <div class="nav-title"><%=Icon("shield")%> Administracion</div>
            <% NavAdmin "dashboard", "admin.asp", "Panel", "chart" %>
            <% NavAdmin "songs", "admin_songs.asp", "Canciones", "music" %>
            <% NavAdmin "artists", "admin_artists.asp", "Artistas", "artist" %>
            <% NavAdmin "albums", "admin_albums.asp", "Albumes", "album" %>
            <% NavAdmin "playlists", "admin_playlists.asp", "Playlists", "playlist" %>
            <% NavAdmin "users", "admin_users.asp", "Usuarios", "users" %>
        </nav>
        <% End If %>

        <div class="sidebar-foot">
            <%=H(APP_NAME)%> v<%=H(APP_VERSION)%><br>
            ASP clasico + SQL Server
        </div>
    </aside>

    <!-- ============================== CONTENIDO ====================== -->
    <div class="main">
        <header class="topbar">
            <button type="button" class="sidebar-toggle" id="pmSidebarToggle" aria-label="Menu">
                <%=Icon("list")%>
            </button>

            <div class="search-box">
                <%=Icon("search")%>
                <input type="search" id="pmSearch" placeholder="Buscar canciones, artistas, albumes..."
                       autocomplete="off" aria-label="Buscar">
            </div>

            <div class="topbar-spacer"></div>

            <div class="user-chip" id="pmUserChip">
                <img src="<%=H(ImgUrl(CurrentUserPhoto(), "user", CurrentUserId(), FirstLetter(CurrentUserName())))%>"
                     alt="<%=H(CurrentUserName())%>">
                <span class="u-text">
                    <span class="u-name"><%=H(CurrentUserName())%></span>
                    <span class="u-role"><% If IsAdmin() Then %>Administrador<% Else %>Usuario<% End If %></span>
                </span>
                <div class="user-menu" id="pmUserMenu">
                    <a href="<%=U("profile.asp")%>"><%=Icon("user")%> Mi perfil</a>
                    <a href="<%=U("home.asp?view=playlists&mine=1")%>" data-nav="playlists"><%=Icon("playlist")%> Mis playlists</a>
                    <% If IsAdmin() Then %>
                    <hr>
                    <a href="<%=U("admin.asp")%>"><%=Icon("shield")%> Panel de administracion</a>
                    <% End If %>
                    <hr>
                    <a href="<%=U("logout.asp")%>"><%=Icon("logout")%> Cerrar sesion</a>
                </div>
            </div>
        </header>

        <main class="view" id="viewRoot" data-view="<%=H(VIEW_NAME_V)%>">
<%=RenderFlash()%>
