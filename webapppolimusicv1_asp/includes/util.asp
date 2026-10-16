<%
' =====================================================================
'  PoliMusic - Utilidades compartidas
' =====================================================================

' ---------------------------------------------------------------------
'  Rutas de la aplicacion (detectadas en tiempo de ejecucion, de modo
'  que la app funciona igual en la raiz del sitio o en un directorio
'  virtual como /webapppolimusicv1_asp/).
'
'  Todas las paginas ASP viven en la raiz de la aplicacion, por lo que la
'  raiz virtual es simplemente el directorio de la pagina en ejecucion.
'  (APPL_MD_PATH no sirve cuando la carpeta es una simple subcarpeta del
'  sitio y no una aplicacion IIS: devolveria siempre "/".)
' ---------------------------------------------------------------------
Function AppBase()
    Dim s
    If Len(APP_VROOT_CACHE) > 0 Then
        AppBase = APP_VROOT_CACHE
        Exit Function
    End If

    s = Request.ServerVariables("SCRIPT_NAME") & ""
    If Len(s) = 0 Then
        s = "/"
    ElseIf InStrRev(s, "/") > 0 Then
        s = Left(s, InStrRev(s, "/"))
    Else
        s = "/"
    End If
    If Len(s) = 0 Then s = "/"

    APP_VROOT_CACHE = s
    AppBase = s
End Function

' Ruta fisica de la raiz de la aplicacion (termina en "\").
Function AppPhysRoot()
    Dim p
    If Len(APP_PHYS_CACHE) > 0 Then
        AppPhysRoot = APP_PHYS_CACHE
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    p = Server.MapPath(AppBase())
    On Error GoTo 0
    If Err.Number <> 0 Or Len(p & "") = 0 Then
        ' Respaldo: raiz del sitio + la ruta virtual de la aplicacion
        p = Server.MapPath("/")
        If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
        p = p & Replace(AppBase(), "/", "\")
    End If
    If Right(p, 1) <> "\" Then p = p & "\"

    APP_PHYS_CACHE = p
    AppPhysRoot = p
End Function

' Construye una URL interna absoluta: U("songs.asp") -> /app/songs.asp
Function U(ByVal sRelative)
    U = AppBase() & sRelative
End Function

' ---------------------------------------------------------------------
'  Codificacion / salida segura
' ---------------------------------------------------------------------
Function H(ByVal v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        H = ""
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&#39;")
    H = s
End Function

' Valor seguro para usar dentro de un atributo entre comillas simples.
Function HJ(ByVal v)
    Dim s
    s = H(v)
    s = Replace(s, "\", "\\")
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    HJ = s
End Function

Function UrlEnc(ByVal v)
    UrlEnc = Server.URLEncode(v & "")
End Function

' Indica si el cuerpo de la peticion es multipart/form-data.
' (Se define aqui porque la usan tanto db.asp como upload.asp.)
Function IsMultipartRequest()
    IsMultipartRequest = (InStr(LCase(Request.ServerVariables("CONTENT_TYPE") & ""), _
                                "multipart/form-data") > 0)
End Function

Function JsonEnc(ByVal v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        JsonEnc = "null"
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonEnc = """" & s & """"
End Function

' ---------------------------------------------------------------------
'  Presentacion de datos
' ---------------------------------------------------------------------
Function FmtDate(ByVal v)
    Dim d
    If IsNull(v) Or Not IsDate(v) Then
        FmtDate = "-"
        Exit Function
    End If
    d = CDate(v)
    FmtDate = Pad2(Day(d)) & "/" & Pad2(Month(d)) & "/" & Year(d)
End Function

Function FmtDateInput(ByVal v)
    Dim d
    If IsNull(v) Or Not IsDate(v) Then
        FmtDateInput = ""
        Exit Function
    End If
    d = CDate(v)
    FmtDateInput = Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d))
End Function

Function FmtNum(ByVal v)
    Dim n
    If IsNull(v) Or Not IsNumeric(v) Then
        FmtNum = "0"
        Exit Function
    End If
    n = CDbl(v)
    FmtNum = CStr(CLng(n))
End Function

Function Trunc(ByVal s, ByVal n)
    Dim t
    t = s & ""
    If Len(t) <= n Then
        Trunc = t
    Else
        Trunc = Left(t, n - 1) & "..."
    End If
End Function

' Nombre completo de un artista (nombre + apellido opcional).
Function ArtistName(ByVal sName, ByVal sLast)
    Dim s
    s = Trim(sName & "")
    If Not IsNull(sLast) Then
        If Len(Trim(sLast & "")) > 0 Then s = s & " " & Trim(sLast)
    End If
    ArtistName = Trim(s)
End Function

' Etiqueta legible de un tipo de usuario / playlist usando TBL_CATALOG.
Function CatalogLabel(ByVal sParent, ByVal nCode, ByVal sDefault)
    Dim v
    v = DbScalar("SELECT TOP 1 [VALUE] FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = " _
        & SqlStr(sParent) & " AND COD_CATALOG = " & SqlStr(CStr(nCode)) & _
        " ORDER BY ID_CATALOG", "")
    If Len(v & "") = 0 Then
        CatalogLabel = sDefault
    Else
        CatalogLabel = v
    End If
End Function

' ---------------------------------------------------------------------
'  Resolucion de imagenes y archivos de audio.
'  En la base de datos las rutas se guardan como "~/carpeta/archivo".
'  Si el archivo no existe en disco se devuelve un marcador SVG generado
'  dinamicamente, de modo que la interfaz nunca muestra una imagen rota.
' ---------------------------------------------------------------------
Function NormalizeAssetRel(ByVal sDbPath)
    Dim p, i, folder, rest
    p = Trim(sDbPath & "")
    If Len(p) = 0 Then
        NormalizeAssetRel = ""
        Exit Function
    End If
    p = Replace(p, "\", "/")
    Do While InStr(p, "~/") > 0
        p = Replace(p, "~/", "")
    Loop
    Do While Left(p, 1) = "/"
        p = Mid(p, 2)
    Loop

    i = InStr(p, "/")
    If i > 0 Then
        folder = LCase(Left(p, i - 1))
        rest = Mid(p, i + 1)
        Select Case folder
            Case "songfiles"     : folder = DIR_SONGS
            Case "userimages"    : folder = DIR_USERS
            Case "playlistimages": folder = DIR_PLAYLISTS
            Case "albumimages"   : folder = DIR_ALBUMS
            Case "artistimages"  : folder = DIR_ARTISTS
            Case "uploads"       : folder = DIR_UPLOADS
            Case Else            : folder = Left(p, i - 1)
        End Select
        NormalizeAssetRel = DIR_ASSETS & "/" & folder & "/" & rest
    Else
        NormalizeAssetRel = DIR_ASSETS & "/" & DIR_UPLOADS & "/" & p
    End If
End Function

Function PlaceholderUrl(ByVal sKind, ByVal sSeed, ByVal sLabel)
    PlaceholderUrl = AppBase() & "placeholder.asp?t=" & UrlEnc(sKind) _
                   & "&s=" & UrlEnc(sSeed & "") & "&l=" & UrlEnc(sLabel & "")
End Function

' Devuelve la URL publica de una imagen (real o marcador generado).
Function ImgUrl(ByVal sDbPath, ByVal sKind, ByVal sSeed, ByVal sLabel)
    Dim rel, phys, fso
    rel = NormalizeAssetRel(sDbPath)
    If Len(rel) = 0 Then
        ImgUrl = PlaceholderUrl(sKind, sSeed, sLabel)
        Exit Function
    End If
    phys = AppPhysRoot() & Replace(rel, "/", "\")
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(phys) Then
        ImgUrl = U(Replace(rel, " ", "%20"))
    Else
        ImgUrl = PlaceholderUrl(sKind, sSeed, sLabel)
    End If
    Set fso = Nothing
End Function

' Devuelve la URL del MP3 solo si el archivo existe (si no, "").
Function AudioUrl(ByVal sDbPath)
    Dim rel, phys, fso
    rel = NormalizeAssetRel(sDbPath)
    If Len(rel) = 0 Then
        AudioUrl = ""
        Exit Function
    End If
    phys = AppPhysRoot() & Replace(rel, "/", "\")
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(phys) Then
        AudioUrl = U(Replace(rel, " ", "%20"))
    Else
        AudioUrl = ""
    End If
    Set fso = Nothing
End Function

' ---------------------------------------------------------------------
'  Mensajes flash (se guardan en sesion y se muestran una sola vez).
' ---------------------------------------------------------------------
Sub SetFlash(ByVal sMessage, ByVal sType)
    Session("flash_msg") = sMessage
    Session("flash_type") = sType
End Sub

Function RenderFlash()
    Dim m, t
    m = Session("flash_msg") & ""
    t = Session("flash_type") & ""
    Session("flash_msg") = ""
    Session("flash_type") = ""
    If Len(m) = 0 Then
        RenderFlash = ""
        Exit Function
    End If
    If Len(t) = 0 Then t = "info"
    RenderFlash = "<div class=""flash flash-" & H(t) & """ role=""status"">" _
        & "<span class=""flash-icon""></span><span>" & H(m) & "</span></div>"
End Function

' ---------------------------------------------------------------------
'  Paginacion
' ---------------------------------------------------------------------
Function PageCount(ByVal nTotal, ByVal nPerPage)
    If nPerPage <= 0 Then nPerPage = 1
    PageCount = Int((nTotal + nPerPage - 1) / nPerPage)
End Function

Function PageOffset(ByVal nPage, ByVal nPerPage)
    If nPage < 1 Then nPage = 1
    PageOffset = (nPage - 1) * nPerPage
End Function

' ---------------------------------------------------------------------
'  Varios
' ---------------------------------------------------------------------
Function FirstLetter(ByVal s)
    Dim t
    t = Trim(s & "")
    If Len(t) = 0 Then
        FirstLetter = "?"
    Else
        FirstLetter = UCase(Left(t, 1))
    End If
End Function

Function NumOr(ByVal v, ByVal nDefault)
    If IsNull(v) Or Not IsNumeric(v) Then
        NumOr = nDefault
    Else
        NumOr = CLng(v)
    End If
End Function

' Acepta unicamente rutas relativas simples (evita redirecciones externas).
Function SafeLocalUrl(ByVal s)
    Dim t
    t = Trim(s & "")
    If Len(t) = 0 Then
        SafeLocalUrl = ""
        Exit Function
    End If
    If InStr(t, "://") > 0 Or InStr(t, "\\") > 0 Or InStr(t, ":") > 0 Then
        SafeLocalUrl = ""
        Exit Function
    End If
    Do While Left(t, 1) = "/" Or Left(t, 1) = "\"
        t = Mid(t, 2)
    Loop
    If InStr(t, "..") > 0 Then
        SafeLocalUrl = ""
        Exit Function
    End If
    SafeLocalUrl = t
End Function

' Valor entero seguro para atributos HTML.
Function AttrInt(ByVal v)
    AttrInt = FmtNum(NumOr(v, 0))
End Function
%>
