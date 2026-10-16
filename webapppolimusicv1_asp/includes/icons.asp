<%
' =====================================================================
'  PoliMusic - Iconografia SVG en linea (sin dependencias externas)
'  Icono(nombre) devuelve el marcado <svg> listo para insertar.
' =====================================================================

Function Icon(ByVal sName)
    Dim o, f, b
    o = "<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""1.9"" stroke-linecap=""round"" stroke-linejoin=""round"" aria-hidden=""true"">"
    f = "<svg viewBox=""0 0 24 24"" fill=""currentColor"" aria-hidden=""true"">"

    Select Case LCase(sName & "")
        Case "home"
            b = "<path d=""M3 10.2 12 3l9 7.2""></path><path d=""M5 9.5V21h14V9.5""></path><path d=""M9.5 21v-6h5v6""></path>"
        Case "music"
            b = "<path d=""M9 18V6l10-2v12""></path><circle cx=""6.5"" cy=""18"" r=""2.6""></circle><circle cx=""16.5"" cy=""16"" r=""2.6""></circle>"
        Case "album"
            b = "<circle cx=""12"" cy=""12"" r=""9""></circle><circle cx=""12"" cy=""12"" r=""2.6""></circle>"
        Case "artist"
            b = "<path d=""M12 15a3.6 3.6 0 0 0 3.6-3.6V6.2a3.6 3.6 0 1 0-7.2 0v5.2A3.6 3.6 0 0 0 12 15z""></path><path d=""M5.5 11.6v.6a6.5 6.5 0 0 0 13 0v-.6""></path><path d=""M12 18.7V22""></path>"
        Case "playlist"
            b = "<path d=""M4 6h11""></path><path d=""M4 11h11""></path><path d=""M4 16h7""></path><path d=""M18 10v8.5""></path><circle cx=""16"" cy=""18.5"" r=""2""></circle>"
        Case "users"
            b = "<circle cx=""9"" cy=""8"" r=""3.4""></circle><path d=""M3 20c0-3.3 2.7-5.6 6-5.6s6 2.3 6 5.6""></path><path d=""M16.5 5.4a3.4 3.4 0 0 1 0 6.6""></path><path d=""M18 20c0-2.3-.8-4.2-2.1-5.3""></path>"
        Case "shield"
            b = "<path d=""M12 3 5 6v5.5c0 4.3 2.9 8.1 7 9.5 4.1-1.4 7-5.2 7-9.5V6l-7-3z""></path><path d=""M9.2 12.2l2 2 3.6-3.9""></path>"
        Case "search"
            b = "<circle cx=""11"" cy=""11"" r=""6.5""></circle><path d=""M16 16l4.5 4.5""></path>"
        Case "play"
            o = f
            b = "<path d=""M7 4.5 19.5 12 7 19.5z""></path>"
        Case "pause"
            o = f
            b = "<rect x=""6.5"" y=""4.5"" width=""4"" height=""15"" rx=""1.2""></rect><rect x=""13.5"" y=""4.5"" width=""4"" height=""15"" rx=""1.2""></rect>"
        Case "skip-back"
            o = f
            b = "<path d=""M18 5.5 8.5 12 18 18.5z""></path><rect x=""5"" y=""5"" width=""2.2"" height=""14"" rx=""1""></rect>"
        Case "skip-forward"
            o = f
            b = "<path d=""M6 5.5 15.5 12 6 18.5z""></path><rect x=""16.8"" y=""5"" width=""2.2"" height=""14"" rx=""1""></rect>"
        Case "shuffle"
            b = "<path d=""M17 4l3 3-3 3""></path><path d=""M17 14l3 3-3 3""></path><path d=""M20 7h-4.2c-1.6 0-3 .9-3.8 2.3L9.4 14.7C8.6 16.1 7.2 17 5.6 17H3""></path><path d=""M20 17h-4.2c-1.6 0-3-.9-3.8-2.3l-.7-1.2""></path><path d=""M3 7h2.6c1.6 0 3 .9 3.8 2.3l.7 1.2""></path>"
        Case "repeat"
            b = "<path d=""M17 3l3 3-3 3""></path><path d=""M20 6H8.5A4.5 4.5 0 0 0 4 10.5V11""></path><path d=""M7 21l-3-3 3-3""></path><path d=""M4 18h11.5A4.5 4.5 0 0 0 20 13.5V13""></path>"
        Case "repeat-one"
            b = "<path d=""M17 3l3 3-3 3""></path><path d=""M20 6H8.5A4.5 4.5 0 0 0 4 10.5V11""></path><path d=""M7 21l-3-3 3-3""></path><path d=""M4 18h11.5A4.5 4.5 0 0 0 20 13.5V13""></path><path d=""M12 11v3""></path>"
        Case "volume"
            b = "<path d=""M4 9.5h3L11 6v12L7 14.5H4z""></path><path d=""M14.5 9.2a4 4 0 0 1 0 5.6""></path><path d=""M17.2 6.8a7.6 7.6 0 0 1 0 10.4""></path>"
        Case "volume-x"
            b = "<path d=""M4 9.5h3L11 6v12L7 14.5H4z""></path><path d=""M15.5 9.5l5 5""></path><path d=""M20.5 9.5l-5 5""></path>"
        Case "plus"
            b = "<path d=""M12 5v14""></path><path d=""M5 12h14""></path>"
        Case "edit"
            b = "<path d=""M4 20h4.2l10-10a2.5 2.5 0 0 0-3.5-3.5l-10 10z""></path><path d=""M14.5 6.5l3 3""></path>"
        Case "trash"
            b = "<path d=""M4 7h16""></path><path d=""M9.5 7V5h5v2""></path><path d=""M6.5 7l1 13h9l1-13""></path><path d=""M10.5 11v6""></path><path d=""M13.5 11v6""></path>"
        Case "logout"
            b = "<path d=""M15 4h3.5A1.5 1.5 0 0 1 20 5.5v13A1.5 1.5 0 0 1 18.5 20H15""></path><path d=""M10 8l-4 4 4 4""></path><path d=""M6 12h9""></path>"
        Case "user"
            b = "<circle cx=""12"" cy=""8.5"" r=""3.6""></circle><path d=""M5 20c0-3.6 3.1-6.1 7-6.1s7 2.5 7 6.1""></path>"
        Case "upload"
            b = "<path d=""M12 16V5""></path><path d=""M8 9l4-4 4 4""></path><path d=""M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2""></path>"
        Case "x"
            b = "<path d=""M6 6l12 12""></path><path d=""M18 6L6 18""></path>"
        Case "check"
            b = "<path d=""M5 12.5l4.5 4.5L19 7""></path>"
        Case "chevron-left"
            b = "<path d=""M14.5 5.5 8 12l6.5 6.5""></path>"
        Case "chevron-right"
            b = "<path d=""M9.5 5.5 16 12l-6.5 6.5""></path>"
        Case "clock"
            b = "<circle cx=""12"" cy=""12"" r=""8.5""></circle><path d=""M12 7.5V12l3 2""></path>"
        Case "chart"
            b = "<path d=""M4 20V10""></path><path d=""M10 20V4""></path><path d=""M16 20v-7""></path><path d=""M22 20H2""></path>"
        Case "sparkle"
            b = "<path d=""M12 3l1.8 5.2L19 10l-5.2 1.8L12 17l-1.8-5.2L5 10l5.2-1.8z""></path><path d=""M18.5 16.5l.8 2.2 2.2.8-2.2.8-.8 2.2-.8-2.2-2.2-.8 2.2-.8z""></path>"
        Case "list"
            b = "<path d=""M8 6h12""></path><path d=""M8 12h12""></path><path d=""M8 18h12""></path><circle cx=""4"" cy=""6"" r=""1.3""></circle><circle cx=""4"" cy=""12"" r=""1.3""></circle><circle cx=""4"" cy=""18"" r=""1.3""></circle>"
        Case "lock"
            b = "<rect x=""5"" y=""10.5"" width=""14"" height=""10"" rx=""2.4""></rect><path d=""M8.5 10.5V8a3.5 3.5 0 0 1 7 0v2.5""></path>"
        Case "globe"
            b = "<circle cx=""12"" cy=""12"" r=""9""></circle><path d=""M3 12h18""></path><path d=""M12 3c2.6 2.6 2.6 15.4 0 18""></path><path d=""M12 3c-2.6 2.6-2.6 15.4 0 18""></path>"
        Case "calendar"
            b = "<rect x=""4"" y=""5.5"" width=""16"" height=""15"" rx=""2.4""></rect><path d=""M4 10h16""></path><path d=""M8.5 3.5v4""></path><path d=""M15.5 3.5v4""></path>"
        Case "mail"
            b = "<rect x=""3"" y=""5.5"" width=""18"" height=""13"" rx=""2.4""></rect><path d=""M4 7.5l8 5.5 8-5.5""></path>"
        Case "image"
            b = "<rect x=""3.5"" y=""4.5"" width=""17"" height=""15"" rx=""2.4""></rect><circle cx=""9"" cy=""10"" r=""1.8""></circle><path d=""M5 17l4.5-4.5 3 3 3-3L19 16""></path>"
        Case "folder"
            b = "<path d=""M3.5 7.5A2 2 0 0 1 5.5 5.5h3.6l2 2.4h7.4a2 2 0 0 1 2 2v7.6a2 2 0 0 1-2 2H5.5a2 2 0 0 1-2-2z""></path>"
        Case Else
            b = "<circle cx=""12"" cy=""12"" r=""9""></circle>"
    End Select

    Icon = o & b & "</svg>"
End Function

' Logo de la aplicacion (nota musical en degradado turquesa).
Function LogoMark()
    LogoMark = "<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""#04211d"" stroke-width=""2"" "_
             & "stroke-linecap=""round"" stroke-linejoin=""round"">"_
             & "<path d=""M9 18V6l10-2v12""></path>"_
             & "<circle cx=""6.5"" cy=""18"" r=""2.6"" fill=""#04211d""></circle>"_
             & "<circle cx=""16.5"" cy=""16"" r=""2.6"" fill=""#04211d""></circle></svg>"
End Function
%>
