<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<%
' =====================================================================
'  PoliMusic - Marcador de imagen generado dinamicamente (SVG)
'  Se usa cuando la ruta guardada en la base de datos no existe en disco,
'  de modo que la interfaz nunca muestra una imagen rota.
'
'  Parametros: t = tipo, s = semilla (para el degradado), l = etiqueta
'  Ejemplo: placeholder.asp?t=album&s=3&l=Rock%20Album
' =====================================================================
Option Explicit
Response.Buffer = True
Response.ContentType = "image/svg+xml"
Response.CharSet = "utf-8"
Response.Expires = 1440

' --- Semilla numerica estable a partir de una cadena -----------------
Function HashNum(ByVal s)
    Dim i, h
    h = 7
    For i = 1 To Len(s)
        h = (h * 31 + Asc(Mid(s, i, 1))) Mod 100003
    Next
    HashNum = h
End Function

' --- Escapado minimo para XML ---------------------------------------
Function XmlEsc(ByVal s)
    Dim r
    r = s & ""
    r = Replace(r, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, """", "&quot;")
    r = Replace(r, "'", "&apos;")
    XmlEsc = r
End Function

Dim sType, sSeed, sLabel, nHash, nHue, nHue2, sLetter, sGlyph

sType  = LCase(Request.QueryString("t") & "")
sSeed  = Request.QueryString("s") & ""
sLabel = Request.QueryString("l") & ""

If Len(sSeed) = 0 Then sSeed = sType & sLabel

nHash = HashNum(sSeed & sType)
nHue  = 150 + (nHash Mod 40)          ' 150..189 -> turquesa / cian
nHue2 = (nHue + 18) Mod 360

sLetter = UCase(Left(Trim(sLabel), 1))
If Len(sLetter) = 0 Or sLetter = "?" Then
    sLetter = ""
End If

' Glifo segun el tipo de recurso
Select Case sType
    Case "artist", "user"
        sGlyph = "<path d=""M256 150a46 46 0 1 0 0 92 46 46 0 0 0 0-92zm0 118c-58 0-105 33-105 74v18h210v-18c0-41-47-74-105-74z""/>"
    Case "playlist"
        sGlyph = "<path d=""M140 170h150v20H140zm0 56h150v20H140zm0 56h96v20h-96zm170-70v150a34 34 0 1 1-20-31V212l-30 8v-22l50-14z""/>"
    Case "album", "song"
        sGlyph = "<path d=""M256 156a100 100 0 1 0 0 200 100 100 0 0 0 0-200zm0 30a70 70 0 1 1 0 140 70 70 0 0 1 0-140zm0 44a26 26 0 1 0 0 52 26 26 0 0 0 0-52z""/>"
    Case Else
        sGlyph = "<path d=""M204 330V186l124-26v142a30 30 0 1 1-18-27V212l-88 19v99a30 30 0 1 1-18-27z""/>"
End Select
%><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 512 512" width="512" height="512" role="img" aria-label="<%=XmlEsc(sLabel)%>">
    <defs>
        <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1">
            <stop offset="0" stop-color="hsl(<%=nHue%>,58%,36%)"/>
            <stop offset="0.55" stop-color="hsl(<%=nHue2%>,52%,24%)"/>
            <stop offset="1" stop-color="hsl(<%=nHue2%>,46%,15%)"/>
        </linearGradient>
    </defs>
    <rect width="512" height="512" fill="url(#bg)"/>
    <circle cx="418" cy="86" r="186" fill="rgba(255,255,255,0.07)"/>
    <circle cx="76" cy="452" r="158" fill="rgba(0,0,0,0.14)"/>
    <g fill="rgba(255,255,255,0.16)"><%=sGlyph%></g>
<% If Len(sLetter) > 0 Then %>
    <text x="256" y="266" text-anchor="middle" dominant-baseline="middle"
          font-family="Segoe UI, Helvetica, Arial, sans-serif" font-size="212"
          font-weight="700" fill="rgba(255,255,255,0.95)"><%=XmlEsc(sLetter)%></text>
<% End If %>
    <rect x="0" y="480" width="512" height="32" fill="rgba(0,0,0,0.22)"/>
</svg>
