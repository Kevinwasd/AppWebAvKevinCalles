<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/sha1.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<!--#include file="includes/icons.asp"-->
<!--#include file="includes/ui.asp"-->
<%
' =====================================================================
'  PoliMusic - Diagnostico del entorno
'  Sin sesion muestra solo las comprobaciones tecnicas (util para
'  detectar problemas antes de poder iniciar sesion).
'  Con sesion de administrador muestra ademas el detalle de los datos.
' =====================================================================
Dim bFull, sProviderUsed, sVersion, rs, i, aTables, aFolders
Dim fso, sPhys, sFile, nOk, nFail, sSha, aChecks, sPath

bFull = IsAdmin()

PAGE_TITLE = "Diagnostico"
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Diagnostico &middot; <%=H(APP_NAME)%></title>
<link rel="icon" href="<%=U("assets/img/favicon.svg")%>" type="image/svg+xml">
<link rel="stylesheet" href="<%=U("assets/css/style.css")%>?v=<%=H(APP_VERSION)%>">
</head>
<body style="background:var(--pm-bg)">
<div style="max-width:1000px;margin:0 auto;padding:34px 22px 60px">

    <div class="section-head">
        <div>
            <h2 class="section-title"><span class="dot"></span>Diagnostico de <%=H(APP_NAME)%></h2>
            <p class="section-sub">Comprobacion de la base de datos, los recursos y el entorno ASP</p>
        </div>
        <a class="btn btn-outline btn-sm" href="<%=U("default.asp")%>">Ir a la aplicacion</a>
    </div>

    <div class="panel">
        <h3 class="panel-title"><%=Icon("shield")%> Entorno de ejecucion</h3>
        <table class="table">
            <tbody>
                <tr><td style="width:280px">Aplicacion</td>
                    <td><b><%=H(APP_NAME)%></b> v<%=H(APP_VERSION)%></td></tr>
                <tr><td>Raiz virtual detectada</td><td><code><%=H(AppBase())%></code></td></tr>
                <tr><td>Raiz fisica detectada</td><td><code><%=H(AppPhysRoot())%></code></td></tr>
                <tr><td>Archivo en ejecucion</td>
                    <td><code><%=H(Request.ServerVariables("SCRIPT_NAME"))%></code></td></tr>
                <tr><td>Servidor web</td>
                    <td><code><%=H(Request.ServerVariables("SERVER_SOFTWARE"))%></code></td></tr>
                <tr><td>Estado de sesion</td>
                    <td><% If IsLoggedIn() Then %>
                        <span class="badge">Sesion activa como <%=H(CurrentUserName())%></span>
                        <% Else %><span class="badge gray">Sin sesion</span><% End If %></td></tr>
            </tbody>
        </table>
    </div>

    <div class="panel">
        <h3 class="panel-title"><%=Icon("check")%> Prueba interna de SHA-1</h3>
        <table class="table">
            <tbody>
                <tr><td style="width:280px">SHA1("abc")</td>
                    <td><code><%=Sha1Hex("abc")%></code>
                    <% If Sha1Hex("abc") = "A9993E364706816ABA3E25717850C26C9CD0D89D" Then %>
                        <span class="badge">Correcto</span>
                    <% Else %><span class="badge danger">Incorrecto</span><% End If %></td></tr>
                <tr><td>SHA1("12345678")</td>
                    <td><code><%=Sha1Hex("12345678")%></code>
                    <% If Sha1Hex("12345678") = "7C222FB2927D828AF22F592134E8932480637C0D" Then %>
                        <span class="badge">Coincide con el script BDD</span>
                    <% Else %><span class="badge danger">No coincide</span><% End If %></td></tr>
            </tbody>
        </table>
    </div>

    <div class="panel">
        <h3 class="panel-title"><%=Icon("album")%> Conexion a SQL Server</h3>
        <%
        On Error Resume Next
        Err.Clear
        Dim oConn
        Set oConn = Server.CreateObject("ADODB.Connection")
        oConn.Open DbConnectionString("MSOLEDBSQL")
        If Err.Number <> 0 Then
            Err.Clear
            oConn.Open DbConnectionString("SQLOLEDB")
            If Err.Number = 0 Then
                sProviderUsed = "SQLOLEDB"
            Else
                sProviderUsed = ""
            End If
        Else
            sProviderUsed = "MSOLEDBSQL"
        End If
        If Len(sProviderUsed) > 0 Then
            sVersion = oConn.Properties("DBMS Version").Value
            oConn.Close
        End If
        Set oConn = Nothing
        On Error GoTo 0
        %>
        <table class="table">
            <tbody>
                <tr><td style="width:280px">Servidor</td><td><code><%=H(DB_SERVER)%></code></td></tr>
                <tr><td>Base de datos</td><td><code><%=H(DB_NAME)%></code></td></tr>
                <tr><td>Usuario</td><td><code><%=H(DB_USER)%></code></td></tr>
                <tr><td>Proveedor OLE DB en uso</td>
                    <td><% If Len(sProviderUsed) > 0 Then %>
                        <code><%=H(sProviderUsed)%></code> <span class="badge">Conectado</span>
                        <% Else %><span class="badge danger">Sin conexion</span><% End If %></td></tr>
                <tr><td>Version del motor</td>
                    <td><code><%=H(IIf2(Len(sVersion) > 0, sVersion, "no disponible"))%></code></td></tr>
            </tbody>
        </table>
        <% If Len(sProviderUsed) = 0 Then %>
        <div class="flash flash-error mt-16"><span class="flash-icon"></span>
            <span>No fue posible conectar. Verifique el servicio de SQL Server, el login
                  <b><%=H(DB_USER)%></b> y la base <b><%=H(DB_NAME)%></b>.</span></div>
        <% End If %>
    </div>

    <% If Len(sProviderUsed) > 0 Then %>
    <div class="panel">
        <h3 class="panel-title"><%=Icon("chart")%> Tablas del modelo</h3>
        <div class="table-wrap">
            <table class="table">
                <thead><tr><th>Tabla</th><th>Filas</th><th>Estado</th></tr></thead>
                <tbody>
                <%
                aTables = Array("TBL_CATALOG", "TBL_SONG", "TBL_USER", "TBL_ALBUM", _
                                "TBL_TRACKLIST", "TBL_ARTIST", "TBL_PLAYER", _
                                "TBL_PLAYLIST", "TBL_PLAYLIST_SONG")
                nOk = 0
                For i = 0 To UBound(aTables)
                    On Error Resume Next
                    Err.Clear
                    sPath = DbScalar("SELECT COUNT(*) FROM " & aTables(i), -1)
                    If Err.Number = 0 Then
                        nOk = nOk + 1
                        Response.Write "<tr><td><code>" & H(aTables(i)) & "</code></td><td>" _
                            & FmtNum(sPath) & "</td><td><span class=""badge"">OK</span></td></tr>"
                    Else
                        Response.Write "<tr><td><code>" & H(aTables(i)) & "</code></td><td>-</td>" _
                            & "<td><span class=""badge danger"">No existe</span></td></tr>"
                    End If
                    On Error GoTo 0
                Next
                %>
                </tbody>
            </table>
        </div>
        <p class="small muted mt-16 mb-0">
            <%=nOk%> de <%=UBound(aTables) + 1%> tablas verificadas correctamente.
        </p>
    </div>
    <% End If %>

    <div class="panel">
        <h3 class="panel-title"><%=Icon("folder")%> Carpetas de recursos</h3>
        <div class="table-wrap">
            <table class="table">
                <thead><tr><th>Carpeta</th><th>Ruta fisica</th><th>Archivos</th><th>Estado</th></tr></thead>
                <tbody>
                <%
                Set fso = Server.CreateObject("Scripting.FileSystemObject")
                aFolders = Array(DIR_SONGS, DIR_USERS, DIR_PLAYLISTS, DIR_ALBUMS, DIR_ARTISTS, DIR_UPLOADS)
                For i = 0 To UBound(aFolders)
                    sPath = AppPhysRoot() & DIR_ASSETS & "\" & aFolders(i)
                    Response.Write "<tr><td><code>assets/" & H(aFolders(i)) & "/</code></td>"
                    Response.Write "<td class=""small muted"">" & H(sPath) & "</td>"
                    If fso.FolderExists(sPath) Then
                        Response.Write "<td>" & fso.GetFolder(sPath).Files.Count & "</td>"
                        Response.Write "<td><span class=""badge"">Disponible</span></td>"
                    Else
                        Response.Write "<td>-</td><td><span class=""badge danger"">Falta</span></td>"
                    End If
                    Response.Write "</tr>"
                Next
                Set fso = Nothing
                %>
                </tbody>
            </table>
        </div>
    </div>

    <% If bFull Then %>
    <div class="panel">
        <h3 class="panel-title"><%=Icon("music")%> Verificacion de los archivos de audio</h3>
        <div class="table-wrap">
            <table class="table">
                <thead><tr><th>ID</th><th>Cancion</th><th>Ruta en la base</th><th>Archivo</th><th>Estado</th></tr></thead>
                <tbody>
                <%
                nOk = 0
                nFail = 0
                Set rs = DbQuery("SELECT ID_SONG, SONG_NAME, SONG_PATH FROM TBL_SONG ORDER BY ID_SONG")
                Do While Not rs.EOF
                    sFile = AppPhysRoot() & Replace(NormalizeAssetRel(rs("SONG_PATH")), "/", "\")
                    Set fso = Server.CreateObject("Scripting.FileSystemObject")
                    Response.Write "<tr><td class=""muted"">#" & CLng(rs("ID_SONG")) & "</td>"
                    Response.Write "<td>" & H(rs("SONG_NAME")) & "</td>"
                    Response.Write "<td class=""small muted"">" & H(rs("SONG_PATH")) & "</td>"
                    Response.Write "<td class=""small muted"">" & H(NormalizeAssetRel(rs("SONG_PATH"))) & "</td>"
                    If fso.FileExists(sFile) Then
                        nOk = nOk + 1
                        Response.Write "<td><span class=""badge"">Encontrado</span></td>"
                    Else
                        nFail = nFail + 1
                        Response.Write "<td><span class=""badge danger"">No encontrado</span></td>"
                    End If
                    Response.Write "</tr>"
                    Set fso = Nothing
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
                %>
                </tbody>
            </table>
        </div>
        <p class="small muted mt-16 mb-0">
            <b><%=nOk%></b> archivos disponibles, <b><%=nFail%></b> no encontrados.
            Los que faltan se reproducen como silencio/error; puede subir el MP3
            desde <a href="<%=U("admin_songs.asp")%>">Administrar canciones</a>.
        </p>
    </div>

    <div class="panel">
        <h3 class="panel-title"><%=Icon("users")%> Cuentas y contrasenas</h3>
        <div class="table-wrap">
            <table class="table">
                <thead><tr><th>ID</th><th>Usuario</th><th>Correo</th><th>Rol</th>
                           <th>Formato de contrasena</th><th>Prueba "12345678"</th></tr></thead>
                <tbody>
                <%
                Set rs = DbQuery("SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, PASSWORD " _
                    & "FROM TBL_USER ORDER BY ID_USER")
                Do While Not rs.EOF
                    Response.Write "<tr><td class=""muted"">#" & CLng(rs("ID_USER")) & "</td>"
                    Response.Write "<td><b>" & H(rs("USERNAME")) & "</b></td>"
                    Response.Write "<td class=""small muted"">" & H(rs("EMAIL")) & "</td>"
                    If CLng(NumOr(rs("USER_TYPE"), 0)) = ROLE_ADMIN Then
                        Response.Write "<td><span class=""badge"">Administrador</span></td>"
                    Else
                        Response.Write "<td><span class=""badge gray"">Usuario</span></td>"
                    End If
                    If Len(rs("PASSWORD") & "") = 40 Then
                        Response.Write "<td><span class=""badge gray"">SHA-1 (40 hex)</span></td>"
                    Else
                        Response.Write "<td><span class=""badge warn"">Texto plano</span></td>"
                    End If
                    If PasswordMatches("12345678", rs("PASSWORD") & "") Then
                        Response.Write "<td><span class=""badge"">Acepta 12345678</span></td>"
                    Else
                        Response.Write "<td><span class=""badge gray"">Otra contrasena</span></td>"
                    End If
                    Response.Write "</tr>"
                    rs.MoveNext
                Loop
                rs.Close
                Set rs = Nothing
                %>
                </tbody>
            </table>
        </div>
    </div>
    <% End If %>

    <div class="panel">
        <h3 class="panel-title"><%=Icon("lock")%> Ajustes recomendados de IIS</h3>
        <p class="small muted">
            Para permitir la carga de archivos MP3 e imagenes desde el panel de
            administracion, el limite de entidad de ASP debe ser mayor que el
            valor por omision (200 KB). El script <code>deploy.ps1</code> ya lo
            configura. Si la carga falla, ejecute en PowerShell como administrador:
        </p>
        <pre style="background:var(--pm-bg-2);padding:14px;border-radius:10px;overflow:auto;font-size:12.5px;color:var(--pm-primary-300)">%windir%\system32\inetsrv\appcmd.exe set config "%APP_POOL_NAME%" -section:asp -limits.maxRequestEntityAllowed:31457280</pre>
        <p class="small muted mb-0">
            Tambien puede dejar el campo <b>Ruta del archivo</b> apuntando a un MP3
            ya copiado en <code>assets/SongFiles/</code> y no subir nada.
        </p>
    </div>

    <p class="small muted">Diagnostico generado el <%=FmtDate(Date())%> a las <%=H(FormatDateTime(Now(), 4))%>.</p>
</div>
</body>
</html>
