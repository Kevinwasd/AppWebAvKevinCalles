<%
' =====================================================================
'  PoliMusic - Enrutador de vistas
'  Las vistas se incluyen estaticamente y se invocan por nombre, de modo
'  que sirven tanto a la pagina completa (home.asp) como a las peticiones
'  parciales (api_view.asp) sin duplicar codigo.
' =====================================================================
%>
<!--#include file="home.asp"-->
<!--#include file="songs.asp"-->
<!--#include file="albums.asp"-->
<!--#include file="album.asp"-->
<!--#include file="artists.asp"-->
<!--#include file="artist.asp"-->
<!--#include file="playlists.asp"-->
<!--#include file="playlist.asp"-->
<!--#include file="search.asp"-->
<%
Sub RenderView(ByVal sView, ByVal sMode)
    VIEW_MODE = sMode
    Select Case LCase(sView & "")
        Case "songs"     : ViewSongs
        Case "albums"    : ViewAlbums
        Case "album"     : ViewAlbum
        Case "artists"   : ViewArtists
        Case "artist"    : ViewArtist
        Case "playlists" : ViewPlaylists
        Case "playlist"  : ViewPlaylist
        Case "search"    : ViewSearch
        Case Else        : ViewHome
    End Select
End Sub

' Devuelve la etiqueta de la vista (para el titulo de la pagina).
Function ViewLabel(ByVal sView)
    Select Case LCase(sView & "")
        Case "songs"     : ViewLabel = "Canciones"
        Case "albums"    : ViewLabel = "Albumes"
        Case "album"     : ViewLabel = "Detalle del album"
        Case "artists"   : ViewLabel = "Artistas"
        Case "artist"    : ViewLabel = "Detalle del artista"
        Case "playlists" : ViewLabel = "Playlists"
        Case "playlist"  : ViewLabel = "Detalle de la playlist"
        Case "search"    : ViewLabel = "Busqueda"
        Case Else        : ViewLabel = "Inicio"
    End Select
End Function
%>
