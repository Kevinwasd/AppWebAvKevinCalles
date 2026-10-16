<%
' =====================================================================
'  PoliMusic - Carga de archivos (multipart/form-data) sin componentes
'  externos: solo ADO (ADODB.Stream / ADODB.Recordset) y FileSystemObject.
'
'  IMPORTANTE: ParseMultipart() debe ejecutarse ANTES de leer
'  Request.Form, porque Request.BinaryRead consume el cuerpo de la
'  peticion una sola vez.
' =====================================================================

Class UploadedFile
    Public FieldName
    Public FileName
    Public ContentType
    Public Data        ' byte-string: 1 caracter = 1 byte
    Public Size
End Class

Dim UP_FIELDS
Dim UP_FILES
Dim UP_PARSED

Set UP_FIELDS = Server.CreateObject("Scripting.Dictionary")
Set UP_FILES  = Server.CreateObject("Scripting.Dictionary")
UP_PARSED = False

' ---------------------------------------------------------------------
'  Conversion binario <-> byte-string
' ---------------------------------------------------------------------

' Byte array (Request.BinaryRead) -> byte-string
Function BytesToString(ByVal binData)
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Fields.Append "b", 201, LenB(binData)     ' 201 = adLongVarBinary
    rs.Open
    rs.AddNew
    rs.Fields("b").AppendChunk binData
    rs.Update
    BytesToString = rs.Fields("b").Value
    rs.Close
    Set rs = Nothing
End Function

' byte-string -> Variant array de enteros 0..255
Function StringToBytes(ByVal sData)
    Dim n, i, a()
    n = Len(sData)
    If n <= 0 Then
        StringToBytes = Array()
        Exit Function
    End If
    ReDim a(n - 1)
    For i = 1 To n
        a(i - 1) = Asc(Mid(sData, i, 1)) And 255
    Next
    StringToBytes = a
End Function

' ---------------------------------------------------------------------
'  Decodificacion UTF-8 de los campos de texto del formulario
' ---------------------------------------------------------------------
Function Utf8ToStr(ByVal sBytes)
    Dim i, n, b, c, d, sOut
    n = Len(sBytes)
    i = 1
    sOut = ""
    Do While i <= n
        b = Asc(Mid(sBytes, i, 1))
        If b < 128 Then
            sOut = sOut & Chr(b)
            i = i + 1
        ElseIf b >= 192 And b < 224 And (i + 1) <= n Then
            c = Asc(Mid(sBytes, i + 1, 1))
            sOut = sOut & Chr(((b - 192) * 64) + (c - 128))
            i = i + 2
        ElseIf b >= 224 And b < 240 And (i + 2) <= n Then
            c = Asc(Mid(sBytes, i + 1, 1))
            d = Asc(Mid(sBytes, i + 2, 1))
            sOut = sOut & Chr(((b - 224) * 4096) + ((c - 128) * 64) + (d - 128))
            i = i + 3
        Else
            sOut = sOut & Chr(b)
            i = i + 1
        End If
    Loop
    Utf8ToStr = sOut
End Function

' ---------------------------------------------------------------------
'  Extrae el valor de una cabecera: name="valor" / filename="valor"
' ---------------------------------------------------------------------
Function HeaderValue(ByVal sHeader, ByVal sKey)
    Dim i, s
    i = InStr(LCase(sHeader), LCase(sKey) & "=""")
    If i = 0 Then
        HeaderValue = ""
        Exit Function
    End If
    s = Mid(sHeader, i + Len(sKey) + 2)
    i = InStr(s, """")
    If i = 0 Then
        HeaderValue = s
    Else
        HeaderValue = Left(s, i - 1)
    End If
End Function

' ---------------------------------------------------------------------
'  Parser principal
' ---------------------------------------------------------------------
Sub ParseMultipart()
    Dim sCt, i, nTotal, sData, sBoundary, sDelim
    Dim aParts, k, sPart, nHead, sHeaders, sBody
    Dim sName, sFile, sCType, oFile

    If UP_PARSED Then Exit Sub
    UP_PARSED = True

    sCt = Request.ServerVariables("CONTENT_TYPE") & ""
    If InStr(LCase(sCt), "multipart/form-data") = 0 Then Exit Sub

    i = InStr(LCase(sCt), "boundary=")
    If i = 0 Then Exit Sub

    sBoundary = Mid(sCt, i + 9)
    sBoundary = Replace(sBoundary, """", "")
    If InStr(sBoundary, ";") > 0 Then
        sBoundary = Left(sBoundary, InStr(sBoundary, ";") - 1)
    End If
    sBoundary = Trim(sBoundary)
    If Len(sBoundary) = 0 Then Exit Sub
    sDelim = "--" & sBoundary

    nTotal = Request.TotalBytes
    If nTotal <= 0 Then Exit Sub

    sData = BytesToString(Request.BinaryRead(nTotal))
    aParts = Split(sData, sDelim)

    ' El primer elemento es el preambulo y el ultimo el delimitador final.
    For k = 1 To UBound(aParts) - 1
        sPart = aParts(k)
        If Left(sPart, 2) = vbCrLf Then sPart = Mid(sPart, 3)
        If Right(sPart, 2) = vbCrLf Then sPart = Left(sPart, Len(sPart) - 2)

        nHead = InStr(sPart, vbCrLf & vbCrLf)
        If nHead > 0 Then
            sHeaders = Left(sPart, nHead - 1)
            sBody = Mid(sPart, nHead + 4)

            sName = HeaderValue(sHeaders, "name")
            sFile = HeaderValue(sHeaders, "filename")
            sCType = ""
            If InStr(LCase(sHeaders), "content-type:") > 0 Then
                sCType = Trim(Mid(sHeaders, InStr(LCase(sHeaders), "content-type:") + 13))
                If InStr(sCType, vbCrLf) > 0 Then sCType = Left(sCType, InStr(sCType, vbCrLf) - 1)
                sCType = Trim(sCType)
            End If

            If Len(sName) > 0 Then
                If Len(sFile) > 0 Then
                    Set oFile = New UploadedFile
                    oFile.FieldName = sName
                    oFile.FileName = SanitizeFileName(sFile)
                    oFile.ContentType = sCType
                    oFile.Data = sBody
                    oFile.Size = Len(sBody)
                    If UP_FILES.Exists(sName) Then UP_FILES.Remove sName
                    UP_FILES.Add sName, oFile
                Else
                    ' Los campos repetidos (grupos de checkbox) se acumulan
                    ' separados por comas, igual que hace Request.Form.
                    If UP_FIELDS.Exists(sName) Then
                        UP_FIELDS(sName) = UP_FIELDS(sName) & "," & Utf8ToStr(sBody)
                    Else
                        UP_FIELDS.Add sName, Utf8ToStr(sBody)
                    End If
                End If
            End If
        End If
    Next
End Sub

' ---------------------------------------------------------------------
'  Acceso a los datos ya analizados
' ---------------------------------------------------------------------

' Valor de un campo de texto (multipart o application/x-www-form-urlencoded)
Function UFField(ByVal sName, ByVal sDefault)
    Dim v
    ParseMultipart()
    v = ""
    If UP_FIELDS.Exists(sName) Then v = UP_FIELDS(sName)
    If Len(v) = 0 And Not IsMultipartRequest() Then v = Request.Form(sName)
    If Len(Trim(v & "")) = 0 Then
        UFField = sDefault
    Else
        UFField = Trim(v)
    End If
End Function

Function UFFieldInt(ByVal sName, ByVal nDefault)
    Dim v
    v = UFField(sName, "")
    If IsNumeric(v) Then
        UFFieldInt = CLng(v)
    Else
        UFFieldInt = nDefault
    End If
End Function

' Devuelve el archivo subido con ese nombre, o Nothing.
Function UFFile(ByVal sName)
    ParseMultipart()
    If UP_FILES.Exists(sName) Then
        Set UFFile = UP_FILES(sName)
    Else
        Set UFFile = Nothing
    End If
End Function

Function HasUpload(ByVal sName)
    Dim f
    Set f = UFFile(sName)
    HasUpload = (Not f Is Nothing) And (f.Size > 0)
End Function

' ---------------------------------------------------------------------
'  Guardado en disco
' ---------------------------------------------------------------------
Function SanitizeFileName(ByVal sName)
    Dim s, i, ch, out
    s = Trim(sName & "")
    If InStr(s, "\") > 0 Then s = Mid(s, InStrRev(s, "\") + 1)
    If InStr(s, "/") > 0 Then s = Mid(s, InStrRev(s, "/") + 1)
    out = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") _
           Or (ch >= "A" And ch <= "Z") Or ch = "." Or ch = "-" Or ch = "_" Then
            out = out & ch
        ElseIf ch = " " Then
            out = out & "_"
        End If
    Next
    If Len(out) = 0 Then out = "archivo"
    SanitizeFileName = out
End Function

Function FileExtension(ByVal sName)
    Dim i
    i = InStrRev(sName & "", ".")
    If i > 0 Then
        FileExtension = LCase(Mid(sName, i + 1))
    Else
        FileExtension = ""
    End If
End Function

' Crea la carpeta destino si no existe.
Sub EnsureAssetFolder(ByVal sFolder)
    Dim fso, phys
    phys = AppPhysRoot() & DIR_ASSETS
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(phys) Then fso.CreateFolder phys
    If Len(sFolder) > 0 Then
        phys = phys & "\" & sFolder
        If Not fso.FolderExists(phys) Then fso.CreateFolder phys
    End If
    Set fso = Nothing
End Sub

' Guarda el archivo subido y devuelve la ruta relativa al estilo de la BD
' ("~/AlbumImages/album_12.png"). Si no hay archivo devuelve "".
Function SaveUploaded(ByVal sField, ByVal sFolder, ByVal sBaseName)
    Dim oFile, ext, rel, phys, stm, ok, fso
    SaveUploaded = ""
    Set oFile = UFFile(sField)
    If oFile Is Nothing Then Exit Function
    If oFile.Size <= 0 Then Exit Function
    If oFile.Size > (MAX_UPLOAD_MB * 1024 * 1024) Then Exit Function

    ext = FileExtension(oFile.FileName)
    If Len(ext) = 0 Then ext = "bin"

    EnsureAssetFolder sFolder
    rel = DIR_ASSETS & "/" & sFolder & "/" & sBaseName & "." & ext
    phys = AppPhysRoot() & Replace(rel, "/", "\")

    ok = False

    ' Via 1: escritura binaria directa con un array de bytes.
    On Error Resume Next
    Err.Clear
    Set stm = Server.CreateObject("ADODB.Stream")
    stm.Type = 1                 ' adTypeBinary
    stm.Open
    stm.Write StringToBytes(oFile.Data)
    stm.SaveToFile phys, 2       ' adSaveCreateOverWrite
    stm.Close
    Set stm = Nothing
    If Err.Number = 0 Then ok = True

    ' Via 2 (respaldo): escritura de texto con juego de caracteres 1:1.
    If Not ok Then
        Err.Clear
        Set stm = Server.CreateObject("ADODB.Stream")
        stm.Type = 2
        stm.Charset = "iso-8859-1"
        stm.Open
        stm.WriteText oFile.Data
        stm.SaveToFile phys, 2
        stm.Close
        Set stm = Nothing
        If Err.Number = 0 Then ok = True
    End If
    On Error GoTo 0

    If Not ok Then Exit Function

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(phys) Then Exit Function
    Set fso = Nothing

    SaveUploaded = "~/" & sFolder & "/" & sBaseName & "." & ext
End Function

' Borra un archivo de recursos indicado con el formato de la BD.
Sub DeleteAsset(ByVal sDbPath)
    Dim rel, phys, fso
    rel = NormalizeAssetRel(sDbPath)
    If Len(rel) = 0 Then Exit Sub
    phys = AppPhysRoot() & Replace(rel, "/", "\")
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(phys) Then fso.DeleteFile phys, True
    Set fso = Nothing
    On Error GoTo 0
End Sub
%>
