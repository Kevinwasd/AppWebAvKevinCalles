<%
' =====================================================================
'  PoliMusic - Autenticacion, sesion y control de acceso por roles
'  Roles: ROLE_ADMIN (1) y ROLE_USER (0)
' =====================================================================

' ---------------------------------------------------------------------
'  Estado de sesion
' ---------------------------------------------------------------------
Function IsLoggedIn()
    IsLoggedIn = (Len(Session("pm_user_id") & "") > 0)
End Function

Function CurrentUserId()
    CurrentUserId = NumOr(Session("pm_user_id"), 0)
End Function

Function CurrentUserName()
    CurrentUserName = Session("pm_user_name") & ""
End Function

Function CurrentUserType()
    CurrentUserType = NumOr(Session("pm_user_type"), ROLE_USER)
End Function

Function IsAdmin()
    IsAdmin = (IsLoggedIn() And CurrentUserType() = ROLE_ADMIN)
End Function

Function CurrentUserPhoto()
    CurrentUserPhoto = Session("pm_user_photo") & ""
End Function

Function CurrentUserEmail()
    CurrentUserEmail = Session("pm_user_email") & ""
End Function

' Establece la sesion a partir de un registro de TBL_USER.
Sub LoginUser(ByVal rsUser)
    Session("pm_user_id")    = CLng(rsUser("ID_USER"))
    Session("pm_user_name")  = rsUser("USERNAME") & ""
    Session("pm_user_type")  = CLng(rsUser("USER_TYPE"))
    Session("pm_user_email") = rsUser("EMAIL") & ""
    Session("pm_user_photo") = rsUser("USER_PHOTO") & ""
End Sub

Sub LogoutUser()
    Session.Contents.RemoveAll()
End Sub

' Recarga los datos de sesion (por ejemplo tras editar el perfil).
Sub RefreshSession()
    Dim rs
    If Not IsLoggedIn() Then Exit Sub
    Set rs = DbQuery("SELECT ID_USER, USERNAME, USER_TYPE, EMAIL, USER_PHOTO "_
        & "FROM TBL_USER WHERE ID_USER = " & CurrentUserId())
    If Not rs.EOF Then
        Session("pm_user_name")  = rs("USERNAME") & ""
        Session("pm_user_type")  = CLng(rs("USER_TYPE"))
        Session("pm_user_email") = rs("EMAIL") & ""
        Session("pm_user_photo") = rs("USER_PHOTO") & ""
    End If
    rs.Close
    Set rs = Nothing
End Sub

' ---------------------------------------------------------------------
'  Guardias de acceso
' ---------------------------------------------------------------------
Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect U("login.asp?next=" & UrlEnc(CurrentRelativeUrl()) & "&e=auth")
    End If
end sub

Sub RequireAdmin()
    RequireLogin()
    If Not IsAdmin() Then
        Response.Redirect U("home.asp?e=denied")
    End If
end sub

' URL relativa de la peticion actual (para volver tras el login).
Function CurrentRelativeUrl()
    Dim s
    s = Request.ServerVariables("SCRIPT_NAME") & ""
    If InStr(s, "/") > 0 Then
        s = Mid(s, InStrRev(s, "/") + 1)
    End If
    If Len(Request.ServerVariables("QUERY_STRING") & "") > 0 Then
        s = s & "?" & Request.ServerVariables("QUERY_STRING")
    End If
    CurrentRelativeUrl = s
End Function

' ---------------------------------------------------------------------
'  Consultas de usuario reutilizables
' ---------------------------------------------------------------------
Function UserById(ByVal nId)
    Set UserById = DbQuery("SELECT TOP 1 * FROM TBL_USER WHERE ID_USER = " & CLng(nId))
End Function

Function UserByLogin(ByVal sLogin)
    Set UserByLogin = DbQuery("SELECT TOP 1 * FROM TBL_USER WHERE USERNAME = " & SqlStr(sLogin))
End Function

Function UsernameExists(ByVal sLogin, ByVal nExceptId)
    Dim sql, n
    sql = "SELECT COUNT(*) FROM TBL_USER WHERE LOWER(USERNAME) = LOWER(" & SqlStr(sLogin) & ")"
    If nExceptId > 0 Then sql = sql & " AND ID_USER <> " & CLng(nExceptId)
    n = CLng(DbScalar(sql, 0))
    UsernameExists = (n > 0)
End Function

Function EmailExists(ByVal sEmail, ByVal nExceptId)
    Dim sql, n
    sql = "SELECT COUNT(*) FROM TBL_USER WHERE LOWER(EMAIL) = LOWER(" & SqlStr(sEmail) & ")"
    If nExceptId > 0 Then sql = sql & " AND ID_USER <> " & CLng(nExceptId)
    n = CLng(DbScalar(sql, 0))
    EmailExists = (n > 0)
End Function
%>
