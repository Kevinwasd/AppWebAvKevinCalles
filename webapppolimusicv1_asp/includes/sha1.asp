<%
' =====================================================================
'  SHA-1 (RFC 3174) implementado en VBScript puro.
'  Se utiliza para almacenar / verificar contrasenas de forma no
'  reversible (TBL_USER.PASSWORD admite texto plano o SHA-1 de 40 hex).
'
'  Notas de implementacion:
'    * VBScript no tiene operadores de desplazamiento ni enteros sin
'      signo de 32 bits, por lo que las sumas se calculan en coma
'      flotante y se reducen modulo 2^32, y los desplazamientos se
'      implementan con multiplicacion / division.
'    * AND / OR / XOR / NOT nativos de VBScript ya operan sobre el
'      patron de bits de 32 bits, que es lo que necesita el algoritmo.
'    * Las constantes se escriben en decimal envueltas en U32() para
'      evitar la ambiguedad de los literales hexadecimales > &H7FFFFFFF.
' =====================================================================

' ---------------------------------------------------------------------
'  Convierte una cadena a su representacion UTF-8 como "byte-string"
'  (1 caracter VBScript = 1 byte, con valores 0..255).
' ---------------------------------------------------------------------
Function Utf8Bytes(ByVal sText)
    Dim i, n, c, sOut
    sOut = ""
    n = Len(sText)
    For i = 1 To n
        c = AscW(Mid(sText, i, 1))
        If c < 0 Then c = c + 65536
        If c < 128 Then
            sOut = sOut & Chr(c)
        ElseIf c < 2048 Then
            sOut = sOut & Chr(192 + Int(c / 64)) & Chr(128 + (c Mod 64))
        Else
            sOut = sOut & Chr(224 + Int(c / 4096)) _
                        & Chr(128 + (Int(c / 64) Mod 64)) _
                        & Chr(128 + (c Mod 64))
        End If
    Next
    Utf8Bytes = sOut
End Function

' ---------------------------------------------------------------------
'  Reduccion modulo 2^32 conservando el patron de 32 bits con signo.
' ---------------------------------------------------------------------
Function U32(ByVal v)
    Dim d
    d = CDbl(v)
    d = d - Int(d / 4294967296.0) * 4294967296.0
    If d < 0 Then d = d + 4294967296.0
    If d >= 2147483648.0 Then d = d - 4294967296.0
    U32 = CLng(d)
End Function

' Valor sin signo (0..2^32-1) de un entero de 32 bits con signo.
Function U32Val(ByVal v)
    If v < 0 Then
        U32Val = CDbl(v) + 4294967296.0
    Else
        U32Val = CDbl(v)
    End If
End Function

Function Add32(ByVal a, ByVal b)
    Add32 = U32(U32Val(a) + U32Val(b))
End Function

Function Shl32(ByVal a, ByVal bits)
    Dim v, i
    v = U32Val(a)
    For i = 1 To bits
        v = v * 2
        If v >= 4294967296.0 Then v = v - 4294967296.0
    Next
    Shl32 = U32(v)
End Function

Function Shr32(ByVal a, ByVal bits)
    Shr32 = U32(Int(U32Val(a) / (2 ^ bits)))
End Function

Function Rotl32(ByVal a, ByVal bits)
    If bits = 0 Then
        Rotl32 = a
    Else
        Rotl32 = (Shl32(a, bits) Or Shr32(a, 32 - bits))
    End If
End Function

' Empaqueta 4 bytes big-endian a partir de un valor de 32 bits.
Function WordToBytes(ByVal v)
    Dim d
    d = U32Val(v)
    WordToBytes = Chr(Int(d / 16777216) Mod 256) _
                & Chr(Int(d / 65536) Mod 256) _
                & Chr(Int(d / 256) Mod 256) _
                & Chr(Int(d) Mod 256)
End Function

' Lee 4 bytes big-endian de un byte-string (posicion base 1).
Function BytesToWord(ByVal sData, ByVal pos)
    Dim b1, b2, b3, b4
    b1 = Asc(Mid(sData, pos, 1))
    b2 = Asc(Mid(sData, pos + 1, 1))
    b3 = Asc(Mid(sData, pos + 2, 1))
    b4 = Asc(Mid(sData, pos + 3, 1))
    BytesToWord = U32(CDbl(b1) * 16777216.0 + CDbl(b2) * 65536.0 _
                    + CDbl(b3) * 256.0 + CDbl(b4))
End Function

' Hexadecimal de 8 digitos (mayusculas) de un valor de 32 bits.
Function Hex32(ByVal v)
    Dim d, s, i, r, hx
    hx = "0123456789ABCDEF"
    d = U32Val(v)
    s = ""
    For i = 1 To 8
        r = CInt(d - Int(d / 16) * 16)
        s = Mid(hx, r + 1, 1) & s
        d = Int(d / 16)
    Next
    Hex32 = s
End Function

' ---------------------------------------------------------------------
'  Sha1Hex : devuelve el resumen SHA-1 en hexadecimal MAYUSCULAS (40).
' ---------------------------------------------------------------------
Function Sha1Hex(ByVal sText)
    Dim sMsg, nLen, nPad, i, k, nBlocks
    Dim h0, h1, h2, h3, h4
    Dim w(79)
    Dim blk, a, b, c, d, e, f, kk, temp

    sMsg = Utf8Bytes(sText)
    nLen = Len(sMsg)

    ' --- Relleno: 0x80 + ceros hasta 56 mod 64 + longitud en bits ---
    sMsg = sMsg & Chr(128)
    nPad = (56 - ((nLen + 1) Mod 64) + 64) Mod 64
    sMsg = sMsg & String(nPad, Chr(0))
    sMsg = sMsg & WordToBytes(0) & WordToBytes(U32(CDbl(nLen) * 8.0))

    h0 = U32(1732584193.0)     ' 67452301
    h1 = U32(4023233417.0)     ' EFCDAB89
    h2 = U32(2562383102.0)     ' 98BADCFE
    h3 = U32(271733878.0)      ' 10325476
    h4 = U32(3285377520.0)     ' C3D2E1F0

    nBlocks = Len(sMsg) \ 64
    For blk = 0 To nBlocks - 1
        For i = 0 To 15
            w(i) = BytesToWord(sMsg, blk * 64 + i * 4 + 1)
        Next
        For i = 16 To 79
            w(i) = Rotl32(w(i - 3) Xor w(i - 8) Xor w(i - 14) Xor w(i - 16), 1)
        Next

        a = h0
        b = h1
        c = h2
        d = h3
        e = h4

        For i = 0 To 79
            If i < 20 Then
                f = (b And c) Or ((Not b) And d)
                kk = U32(1518500249.0)     ' 5A827999
            ElseIf i < 40 Then
                f = b Xor c Xor d
                kk = U32(1859775393.0)     ' 6ED9EBA1
            ElseIf i < 60 Then
                f = (b And c) Or (b And d) Or (c And d)
                kk = U32(2400959708.0)     ' 8F1BBCDC
            Else
                f = b Xor c Xor d
                kk = U32(3395469782.0)     ' CA62C1D6
            End If

            temp = Add32(Add32(Add32(Rotl32(a, 5), f), Add32(e, kk)), w(i))
            e = d
            d = c
            c = Rotl32(b, 30)
            b = a
            a = temp
        Next

        h0 = Add32(h0, a)
        h1 = Add32(h1, b)
        h2 = Add32(h2, c)
        h3 = Add32(h3, d)
        h4 = Add32(h4, e)
    Next

    Sha1Hex = Hex32(h0) & Hex32(h1) & Hex32(h2) & Hex32(h3) & Hex32(h4)
End Function

' ---------------------------------------------------------------------
'  Verifica una contrasena contra el valor almacenado, aceptando tanto
'  texto plano (datos de ejemplo del script BDD) como SHA-1 hexadecimal.
' ---------------------------------------------------------------------
Function PasswordMatches(ByVal sInput, ByVal sStored)
    Dim a, b
    If IsNull(sInput) Or IsNull(sStored) Then
        PasswordMatches = False
        Exit Function
    End If
    a = Trim(sStored & "")
    b = Trim(sInput & "")
    If Len(a) = 0 Then
        PasswordMatches = False
        Exit Function
    End If
    If UCase(a) = Sha1Hex(b) Then
        PasswordMatches = True
    ElseIf a = b Then
        PasswordMatches = True
    Else
        PasswordMatches = False
    End If
End Function
%>
