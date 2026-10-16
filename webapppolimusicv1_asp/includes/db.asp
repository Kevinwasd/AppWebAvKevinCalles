<%
' =====================================================================
'  PoliMusic - Capa de acceso a datos (ADO / OLE DB)
'  Todos los accesos a SQL Server pasan por estas funciones.
' =====================================================================

' ---------------------------------------------------------------------
'  Cadena de conexion traducida al formato nativo de OLE DB.
'  Se respetan los parametros Encrypt y TrustServerCertificate
'  solicitados en la configuracion.
' ---------------------------------------------------------------------
Function DbConnectionString(ByVal sProvider)
    Dim cs
    cs = "Provider=" & sProvider & ";" _
       & "Data Source=" & DB_SERVER & ";" _
       & "Initial Catalog=" & DB_NAME & ";" _
       & "User ID=" & DB_USER & ";" _
       & "Password=" & DB_PASSWORD & ";"

    Select Case sProvider
        Case "MSOLEDBSQL", "MSOLEDBSQL19"
            ' Encrypt=Optional equivale a "preferir cifrado" y
            ' TrustServerCertificate=Yes acepta el certificado propio.
            cs = cs & "Encrypt=Optional;TrustServerCertificate=Yes;"
        Case "SQLNCLI11", "SQLNCLI10"
            cs = cs & "Encrypt=No;TrustServerCertificate=Yes;"
        Case Else
            cs = cs & "Persist Security Info=False;"
    End Select

    DbConnectionString = cs & "Persist Security Info=False;Application Name=" & APP_NAME & ";"
End Function

' ---------------------------------------------------------------------
'  Abre una conexion probando los proveedores OLE DB disponibles.
' ---------------------------------------------------------------------
Function DbOpen()
    Dim providers, i, c, ok
    providers = Array("MSOLEDBSQL", "MSOLEDBSQL19", "SQLNCLI11", "SQLOLEDB")

    For i = 0 To UBound(providers)
        Set c = Server.CreateObject("ADODB.Connection")
        c.ConnectionTimeout = 10
        c.CommandTimeout = 60
        ok = False

        On Error Resume Next
        Err.Clear
        c.Open DbConnectionString(providers(i))
        If Err.Number = 0 Then ok = True
        On Error GoTo 0

        If ok Then
            Set DbOpen = c
            Exit Function
        End If

        On Error Resume Next
        c.Close
        On Error GoTo 0
        Set c = Nothing
    Next

    ' Ningun proveedor funciono
    DbFatalError "No fue posible conectar con SQL Server (" & DB_SERVER & _
                 ", base de datos " & DB_NAME & ")."
End Function

' Muestra un error fatal de base de datos y detiene la pagina.
Sub DbFatalError(ByVal sMessage)
    Response.Write "<div style=""font-family:Segoe UI,Arial,sans-serif;max-width:760px;"_
        & "margin:60px auto;padding:28px 32px;border:1px solid #f0b4b4;border-radius:14px;"_
        & "background:#fff5f5;color:#7a1c1c;"">"
    Response.Write "<h2 style=""margin:0 0 10px"">Error de base de datos</h2>"
    Response.Write "<p style=""margin:0 0 14px"">" & sMessage & "</p>"
    Response.Write "<p style=""margin:0;font-size:13px;color:#96595a"">Verifique que el "_
        & "servicio SQL Server este activo, que el login <b>" & DB_USER & "</b> exista y "_
        & "que la base <b>" & DB_NAME & "</b> haya sido creada con el script "_
        & "<b>BDD_PoliMusic.sql</b>.</p></div>"
    Response.End
End Sub

' ---------------------------------------------------------------------
'  Ejecuta una consulta SELECT y devuelve un Recordset desconectado
'  (cursor de cliente), de modo que la conexion se libera enseguida.
' ---------------------------------------------------------------------
Function DbQuery(ByVal sql)
    Dim c, rs
    Set c = DbOpen()
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3          ' adUseClient
    rs.Open sql, c, 3, 1           ' adOpenStatic, adLockReadOnly
    Set rs.ActiveConnection = Nothing
    c.Close
    Set c = Nothing
    Set DbQuery = rs
End Function

' Devuelve el primer campo de la primera fila (o sDefault si no hay filas).
Function DbScalar(ByVal sql, ByVal sDefault)
    Dim rs, v
    Set rs = DbQuery(sql)
    v = sDefault
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0).Value) Then v = rs.Fields(0).Value
    End If
    rs.Close
    Set rs = Nothing
    DbScalar = v
End Function

' Ejecuta una sentencia sin resultados (INSERT / UPDATE / DELETE).
Sub DbExec(ByVal sql)
    Dim c
    Set c = DbOpen()
    c.Execute sql
    c.Close
    Set c = Nothing
End Sub

' Ejecuta un INSERT y devuelve el IDENTITY generado (0 si falla).
' SET NOCOUNT ON evita los mensajes de "filas afectadas", de modo que el
' primer conjunto de resultados del lote es siempre el SELECT final.
Function DbInsert(ByVal sql)
    Dim rs, v
    v = 0
    On Error Resume Next
    Err.Clear
    Set rs = DbQuery("SET NOCOUNT ON; " & sql & " SELECT CAST(SCOPE_IDENTITY() AS int) AS NEW_ID;")
    If Err.Number = 0 Then
        If Not rs.EOF Then
            If Not IsNull(rs.Fields(0).Value) Then v = CLng(rs.Fields(0).Value)
        End If
        rs.Close
    End If
    On Error GoTo 0
    DbInsert = v
End Function

' ---------------------------------------------------------------------
'  Escapes de seguridad: TODA entrada de usuario se normaliza antes de
'  concatenarse en SQL. Evita inyeccion de SQL por comillas y por tipos.
' ---------------------------------------------------------------------
Function SqlStr(ByVal v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        SqlStr = "NULL"
        Exit Function
    End If
    s = CStr(v)
    s = Replace(s, vbNullChar, "")
    s = Replace(s, "'", "''")
    SqlStr = "N'" & s & "'"
End Function

' Igual que SqlStr pero la cadena vacia se guarda como cadena vacia.
Function SqlStrE(ByVal v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        SqlStrE = "''"
        Exit Function
    End If
    s = Replace(CStr(v), "'", "''")
    SqlStrE = "N'" & s & "'"
End Function

Function SqlInt(ByVal v, ByVal nDefault)
    If IsNumeric(v) Then
        SqlInt = CLng(v)
    Else
        SqlInt = nDefault
    End If
End Function

' Entero o NULL (util para claves foraneas opcionales).
Function SqlIntNull(ByVal v)
    Dim n
    If Not IsNumeric(v) Then
        SqlIntNull = "NULL"
        Exit Function
    End If
    n = CLng(v)
    If n <= 0 Then
        SqlIntNull = "NULL"
    Else
        SqlIntNull = n
    End If
End Function

Function SqlDate(ByVal v)
    Dim d
    If Not IsDate(v) Then
        SqlDate = "NULL"
        Exit Function
    End If
    d = CDate(v)
    SqlDate = "'" & Year(d) & "-" & Pad2(Month(d)) & "-" & Pad2(Day(d)) & "'"
End Function

Function SqlBit(ByVal v)
    If IsNumeric(v) Then
        If CLng(v) <> 0 Then
            SqlBit = 1
        Else
            SqlBit = 0
        End If
    Else
        SqlBit = 0
    End If
End Function

Function Pad2(ByVal n)
    If n < 10 Then
        Pad2 = "0" & n
    Else
        Pad2 = CStr(n)
    End If
End Function

' ---------------------------------------------------------------------
'  Validacion de parametros de entrada.
' ---------------------------------------------------------------------

' Lee un parametro de querystring o de formulario, sin espacios extremos.
' La querystring se consulta primero: en las peticiones multipart no se
' puede acceder a Request.Form despues de Request.BinaryRead.
Function Param(ByVal sName, ByVal sDefault)
    Dim v
    v = Request.QueryString(sName)
    If Len(Trim(v & "")) = 0 Then
        If Not IsMultipartRequest() Then
            v = Request.Form(sName)
        End If
    End If
    If Len(Trim(v & "")) = 0 Then
        Param = sDefault
    Else
        Param = Trim(v)
    End If
End Function

' Lee un parametro numerico entero.
Function ParamInt(ByVal sName, ByVal nDefault)
    Dim v
    v = Param(sName, "")
    If IsNumeric(v) Then
        ParamInt = CLng(v)
    Else
        ParamInt = nDefault
    End If
End Function

' Lee un parametro limitando su longitud (evita entradas gigantes).
Function ParamMax(ByVal sName, ByVal nMax)
    Dim v
    v = Param(sName, "")
    If Len(v) > nMax Then v = Left(v, nMax)
    ParamMax = v
End Function
%>
