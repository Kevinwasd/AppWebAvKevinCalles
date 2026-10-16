<%
' =====================================================================
'  PoliMusic - Acciones enviadas por formularios (POST)
'  Se procesan antes de renderizar la pagina y finalizan con una
'  redireccion (patron POST/Redirect/GET) para evitar reenvios.
' =====================================================================
Sub HandlePostActions()
    Dim sAction, nId, nPl, nSong, sMsg, nRes

    If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then Exit Sub

    sAction = LCase(UFField("action", ""))
    If Len(sAction) = 0 Then Exit Sub

    If Not IsLoggedIn() Then
        Response.Redirect U("login.asp?e=auth")
    End If

    Select Case sAction

        ' ------------------- Crear / actualizar playlist ---------------
        Case "playlist_save"
            If UFField("do_delete", "") = "1" Then
                nId = UFFieldInt("id", 0)
                If nId > 0 Then
                    PlaylistDelete nId
                    SetFlash "La playlist fue eliminada correctamente.", "success"
                End If
                Response.Redirect U("home.asp?view=playlists&mine=1")
            End If

            nRes = PlaylistSaveFromRequest(sMsg)
            If Len(sMsg) > 0 Then
                If nRes > 0 Then
                    SetFlash sMsg, "success"
                Else
                    SetFlash sMsg, "error"
                End If
            End If
            If nRes > 0 Then
                Response.Redirect U("home.asp?view=playlist&id=" & nRes)
            Else
                Response.Redirect U("home.asp?view=playlists&mine=1")
            End If

        ' --------------------------- Eliminar playlist -----------------
        Case "playlist_delete"
            nId = UFFieldInt("id", 0)
            If nId > 0 Then
                PlaylistDelete nId
                SetFlash "La playlist fue eliminada correctamente.", "success"
            End If
            Response.Redirect U("home.asp?view=playlists&mine=1")

        ' ---------------------- Agregar cancion ------------------------
        Case "playlist_add"
            nPl = UFFieldInt("playlist", 0)
            nSong = UFFieldInt("song", 0)
            sMsg = PlaylistAddSong(nPl, nSong)
            If Len(sMsg) = 0 Then
                SetFlash "La cancion fue agregada a la playlist.", "success"
            Else
                SetFlash sMsg, "error"
            End If
            Response.Redirect U("home.asp?view=playlist&id=" & nPl)

        ' ----------------------- Quitar cancion ------------------------
        Case "playlist_remove"
            nPl = UFFieldInt("playlist", 0)
            nSong = UFFieldInt("song", 0)
            PlaylistRemoveSong nPl, nSong
            SetFlash "La cancion fue quitada de la playlist.", "success"
            Response.Redirect U("home.asp?view=playlist&id=" & nPl)

        ' ------------------------- Perfil propio -----------------------
        Case "profile_save"
            nRes = ProfileSaveFromRequest(sMsg)
            If nRes > 0 Then
                SetFlash sMsg, "success"
            Else
                SetFlash sMsg, "error"
            End If
            Response.Redirect U("profile.asp")

    End Select
End Sub

' ---------------------------------------------------------------------
'  Actualizacion del perfil del usuario conectado.
'  Devuelve 1 si la operacion fue correcta, 0 en caso contrario.
' ---------------------------------------------------------------------
Function ProfileSaveFromRequest(ByRef sMsg)
    Dim sUser, sEmail, sBirth, sPass, sPass2, sPhoto, nId, nDup

    nId    = CurrentUserId()
    sUser  = Left(UFField("username", ""), 50)
    sEmail = Left(UFField("email", ""), 50)
    sBirth = UFField("birthday", "")
    sPass  = UFField("password", "")
    sPass2 = UFField("password2", "")

    sMsg = ""

    If Len(sUser) < 4 Then
        sMsg = "El nombre de usuario debe tener al menos 4 caracteres."
        ProfileSaveFromRequest = 0
        Exit Function
    End If
    If Len(sEmail) < 6 Or InStr(sEmail, "@") = 0 Then
        sMsg = "Debe indicar un correo electronico valido."
        ProfileSaveFromRequest = 0
        Exit Function
    End If
    If UsernameExists(sUser, nId) Then
        sMsg = "El nombre de usuario ya esta en uso por otra cuenta."
        ProfileSaveFromRequest = 0
        Exit Function
    End If
    If EmailExists(sEmail, nId) Then
        sMsg = "El correo electronico ya esta en uso por otra cuenta."
        ProfileSaveFromRequest = 0
        Exit Function
    End If
    If Len(sBirth) > 0 And Not IsDate(sBirth) Then
        sMsg = "La fecha de nacimiento no es valida."
        ProfileSaveFromRequest = 0
        Exit Function
    End If
    If Len(sPass) > 0 Then
        If Len(sPass) < 6 Then
            sMsg = "La nueva contrasena debe tener al menos 6 caracteres."
            ProfileSaveFromRequest = 0
            Exit Function
        End If
        If sPass <> sPass2 Then
            sMsg = "Las contrasenas no coinciden."
            ProfileSaveFromRequest = 0
            Exit Function
        End If
    End If

    DbExec "UPDATE TBL_USER SET USERNAME = " & SqlStr(sUser) _
        & ", EMAIL = " & SqlStr(sEmail) _
        & ", BIRTHDAY = " & SqlDate(sBirth) _
        & " WHERE ID_USER = " & nId

    If Len(sPass) > 0 Then
        DbExec "UPDATE TBL_USER SET PASSWORD = " & SqlStr(Sha1Hex(sPass)) _
            & " WHERE ID_USER = " & nId
    End If

    If HasUpload("photo") Then
        sPhoto = SaveUploaded("photo", DIR_USERS, "user_" & nId)
        If Len(sPhoto) > 0 Then
            DbExec "UPDATE TBL_USER SET USER_PHOTO = " & SqlStr(sPhoto) _
                & " WHERE ID_USER = " & nId
        End If
    End If

    RefreshSession
    sMsg = "Sus datos fueron actualizados correctamente."
    ProfileSaveFromRequest = 1
End Function
%>
