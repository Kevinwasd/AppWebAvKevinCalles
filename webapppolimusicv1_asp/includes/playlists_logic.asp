<%
' =====================================================================
'  PoliMusic - Reglas de negocio de las playlists
'  Compartidas por las vistas, los formularios y los servicios JSON.
' =====================================================================

Function PlaylistExists(ByVal nId)
    PlaylistExists = (CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & CLng(nId), 0)) > 0)
End Function

Function PlaylistOwner(ByVal nId)
    PlaylistOwner = CLng(DbScalar("SELECT ISNULL(MIN(ID_USER),0) FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & CLng(nId), 0))
End Function

' El propietario de la playlist y el administrador pueden modificarla.
Function PlaylistCanEdit(ByVal nId)
    Dim nOwner
    If Not IsLoggedIn() Then
        PlaylistCanEdit = False
        Exit Function
    End If
    If IsAdmin() Then
        PlaylistCanEdit = True
        Exit Function
    End If
    nOwner = PlaylistOwner(nId)
    PlaylistCanEdit = (nOwner > 0 And nOwner = CurrentUserId())
End Function

' Agrega una cancion. Devuelve "" si todo fue bien o el mensaje de error.
Function PlaylistAddSong(ByVal nPlaylist, ByVal nSong)
    Dim n
    PlaylistAddSong = ""
    If Not PlaylistExists(nPlaylist) Then
        PlaylistAddSong = "La playlist no existe."
        Exit Function
    End If
    If Not PlaylistCanEdit(nPlaylist) Then
        PlaylistAddSong = "No tiene permiso para modificar esta playlist."
        Exit Function
    End If
    If CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = " & CLng(nSong), 0)) = 0 Then
        PlaylistAddSong = "La cancion seleccionada no existe."
        Exit Function
    End If
    n = CLng(DbScalar("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & CLng(nPlaylist) _
        & " AND ID_SONG = " & CLng(nSong), 0))
    If n > 0 Then
        PlaylistAddSong = "La cancion ya forma parte de la playlist."
        Exit Function
    End If
    DbExec "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (" _
        & CLng(nPlaylist) & ", " & CLng(nSong) & ")"
    PlaylistAddSong = ""
End Function

Sub PlaylistRemoveSong(ByVal nPlaylist, ByVal nSong)
    If Not PlaylistCanEdit(nPlaylist) Then Exit Sub
    DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & CLng(nPlaylist) _
        & " AND ID_SONG = " & CLng(nSong)
End Sub

Sub PlaylistDelete(ByVal nId)
    Dim sPhoto
    If Not PlaylistCanEdit(nId) Then Exit Sub
    sPhoto = DbScalar("SELECT ISNULL(MIN(PHOTO),'') FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & CLng(nId), "")
    DbExec "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = " & CLng(nId)
    DbExec "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = " & CLng(nId)
    If Len(sPhoto & "") > 0 Then DeleteAsset sPhoto
End Sub

' Crea o actualiza una playlist a partir de los datos del formulario.
' Devuelve el id afectado y deja el mensaje en sMsg.
Function PlaylistSaveFromRequest(ByRef sMsg)
    Dim nId, sTitle, nType, sPhoto, nNew
    nId = UFFieldInt("id", 0)
    sTitle = Left(UFField("title", ""), 50)
    nType = UFFieldInt("type", PLAYLIST_PUBLIC)
    If nType <> PLAYLIST_PRIVATE Then nType = PLAYLIST_PUBLIC

    sMsg = ""
    If Len(sTitle) = 0 Then
        sMsg = "El titulo de la playlist es obligatorio."
        PlaylistSaveFromRequest = 0
        Exit Function
    End If

    If nId > 0 Then
        If Not PlaylistCanEdit(nId) Then
            sMsg = "No tiene permiso para modificar esta playlist."
            PlaylistSaveFromRequest = 0
            Exit Function
        End If
        DbExec "UPDATE TBL_PLAYLIST SET TITLE = " & SqlStr(sTitle) & ", TYPE = " & nType _
            & " WHERE ID_PLAYLIST = " & nId
        If HasUpload("photo") Then
            sPhoto = SaveUploaded("photo", DIR_PLAYLISTS, "playlist_" & nId)
            If Len(sPhoto) > 0 Then
                DbExec "UPDATE TBL_PLAYLIST SET PHOTO = " & SqlStr(sPhoto) _
                    & " WHERE ID_PLAYLIST = " & nId
            End If
        End If
        sMsg = "Playlist actualizada correctamente."
        PlaylistSaveFromRequest = nId
    Else
        nNew = DbInsert("INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (" _
            & CurrentUserId() & ", " & SqlStr(sTitle) & ", GETDATE(), " & nType & ", " _
            & SqlStr("~/PlaylistImages/default.jpg") & ")")
        If nNew <= 0 Then
            sMsg = "No fue posible crear la playlist."
            PlaylistSaveFromRequest = 0
            Exit Function
        End If
        If HasUpload("photo") Then
            sPhoto = SaveUploaded("photo", DIR_PLAYLISTS, "playlist_" & nNew)
            If Len(sPhoto) > 0 Then
                DbExec "UPDATE TBL_PLAYLIST SET PHOTO = " & SqlStr(sPhoto) _
                    & " WHERE ID_PLAYLIST = " & nNew
            End If
        End If
        sMsg = "Playlist creada correctamente."
        PlaylistSaveFromRequest = nNew
    End If
End Function
%>
