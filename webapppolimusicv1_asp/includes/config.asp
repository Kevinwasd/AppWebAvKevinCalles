<%
' =====================================================================
'  PoliMusic - Configuracion global
'  Tecnologia : ASP Clasico (VBScript) + SQL Server
'  Autor      : Equipo PoliMusic
' =====================================================================
Option Explicit

' ------------------------------ Aplicacion ---------------------------
Const APP_NAME     = "PoliMusic"
Const APP_VERSION  = "1.0.0"
Const APP_TAGLINE  = "Tu musica, sin limites"

' -------------------------- Conexion SQL Server ----------------------
' Cadena solicitada:
'   Server=localhost; Database=BDD_PoliMusic; User Id=usr_polimusic;
'   Password=Politecnica1; Encrypt=true; TrustServerCertificate=true
' ADO (ASP clasico) no acepta el formato "Server=..;Database=..", por eso
' se traduce a los nombres nativos de OLE DB respetando Encrypt y
' TrustServerCertificate. Ver includes/db.asp -> DbConnectionString().
Const DB_SERVER     = "localhost"
Const DB_NAME       = "BDD_PoliMusic"
Const DB_USER       = "usr_polimusic"
Const DB_PASSWORD   = "Politecnica1"
Const DB_ENCRYPT    = "true"
Const DB_TRUST_CERT = "true"

' -------------------------- Roles de usuario -------------------------
' TBL_CATALOG : COD_CATALOG_PARENT = '1' -> USER TYPE
Const ROLE_USER    = 0      ' Usuario normal
Const ROLE_ADMIN   = 1      ' Administrador

' -------------------------- Tipos de playlist ------------------------
' TBL_CATALOG : COD_CATALOG_PARENT = '2' -> PLAYLIST TYPE
Const PLAYLIST_PUBLIC  = 0
Const PLAYLIST_PRIVATE = 1

' -------------------------- Carpetas de recursos ---------------------
Const DIR_ASSETS    = "assets"
Const DIR_SONGS     = "SongFiles"
Const DIR_USERS     = "UserImages"
Const DIR_PLAYLISTS = "PlaylistImages"
Const DIR_ALBUMS    = "AlbumImages"
Const DIR_ARTISTS   = "ArtistImages"
Const DIR_UPLOADS   = "uploads"

' -------------------------- Limites y paginacion ---------------------
Const PAGE_SIZE     = 12
Const MAX_UPLOAD_MB = 20

' -------------------------- Estado interno ---------------------------
Dim APP_VROOT_CACHE
Dim APP_PHYS_CACHE
APP_VROOT_CACHE = ""
APP_PHYS_CACHE  = ""

' Variables de presentacion que cada pagina puede asignar antes de
' incluir layout_top.asp (se declaran aqui para que Option Explicit
' funcione en todos los archivos).
Dim PAGE_TITLE, VIEW_NAME, ACTIVE_ADMIN, VIEW_MODE
PAGE_TITLE   = ""
VIEW_NAME    = ""
ACTIVE_ADMIN = ""
VIEW_MODE    = "full"

Response.Buffer  = True
Response.CharSet = "utf-8"

' El codepage de sesion se fija de forma tolerante: si el servidor no lo
' permite, la aplicacion sigue funcionando (todo el marcado es ASCII con
' entidades HTML, por lo que la codificacion no afecta a la interfaz).
On Error Resume Next
Session.CodePage = 65001
On Error GoTo 0
%>
