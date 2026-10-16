<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="includes/config.asp"-->
<!--#include file="includes/db.asp"-->
<!--#include file="includes/util.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' =====================================================================
'  PoliMusic - Registro de reproducciones
'  POST id=<ID_SONG>  ->  { "ok": true, "id": 3, "plays": 12 }
' =====================================================================
Dim nId, nPlays, bOk, sOk

Response.ContentType = "application/json"
Response.CacheControl = "no-cache"
Response.CharSet = "utf-8"

nId = ParamInt("id", 0)
bOk = False
nPlays = 0

If IsLoggedIn() And nId > 0 Then
    If CLng(DbScalar("SELECT COUNT(*) FROM TBL_SONG WHERE ID_SONG = " & nId, 0)) > 0 Then
        DbExec "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = " & nId
        bOk = True
    End If
End If

If nId > 0 Then
    nPlays = CLng(DbScalar("SELECT ISNULL(PLAYS,0) FROM TBL_SONG WHERE ID_SONG = " & nId, 0))
End If

If bOk Then
    sOk = "true"
Else
    sOk = "false"
End If

Response.Write "{""ok"":" & sOk & ",""id"":" & nId & ",""plays"":" & nPlays & "}"
%>
