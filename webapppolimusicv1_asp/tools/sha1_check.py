"""
Transliteracion linea a linea de includes/sha1.asp a Python.

Objetivo: validar el DISENO del algoritmo (relleno, expansion del mensaje,
80 rondas, mezcla final). Si esta transliteracion coincide con hashlib,
la logica escrita en VBScript es correcta y solo queda el riesgo sintactico
del lenguaje (cubierto por tools/vbslint.py).
"""
import hashlib


def U32(v):
    d = float(v)
    d = d - int(d / 4294967296.0) * 4294967296.0
    if d < 0:
        d += 4294967296.0
    if d >= 2147483648.0:
        d -= 4294967296.0
    return int(d)


def U32Val(v):
    return float(v) + 4294967296.0 if v < 0 else float(v)


def i32(v):
    """VBScript Long: patron de 32 bits con signo."""
    v &= 0xFFFFFFFF
    return v - 0x100000000 if v >= 0x80000000 else v


def Add32(a, b):
    return U32(U32Val(a) + U32Val(b))


def Shl32(a, bits):
    v = U32Val(a)
    for _ in range(bits):
        v = v * 2
        if v >= 4294967296.0:
            v -= 4294967296.0
    return U32(v)


def Shr32(a, bits):
    return U32(int(U32Val(a) / (2 ** bits)))


def Rotl32(a, bits):
    if bits == 0:
        return a
    return i32(Shl32(a, bits) | Shr32(a, 32 - bits))


def WordToBytes(v):
    d = U32Val(v)
    return bytes([int(d / 16777216) % 256, int(d / 65536) % 256,
                  int(d / 256) % 256, int(d) % 256])


def BytesToWord(data, pos):
    b1, b2, b3, b4 = data[pos - 1], data[pos], data[pos + 1], data[pos + 2]
    return U32(b1 * 16777216.0 + b2 * 65536.0 + b3 * 256.0 + b4)


def Hex32(v):
    d = U32Val(v)
    s = ""
    for _ in range(8):
        s = "0123456789ABCDEF"[int(d) % 16] + s
        d = int(d / 16)
    return s


def Utf8Bytes(text):
    return text.encode("utf-8")


def Sha1Hex(text):
    sMsg = bytearray(Utf8Bytes(text))
    nLen = len(sMsg)

    sMsg.append(128)
    nPad = (56 - ((nLen + 1) % 64) + 64) % 64
    sMsg.extend(bytes(nPad))
    sMsg.extend(WordToBytes(0))
    sMsg.extend(WordToBytes(U32(float(nLen) * 8.0)))

    h0 = U32(1732584193.0)
    h1 = U32(4023233417.0)
    h2 = U32(2562383102.0)
    h3 = U32(271733878.0)
    h4 = U32(3285377520.0)

    nBlocks = len(sMsg) // 64
    for blk in range(nBlocks):
        w = [0] * 80
        for i in range(16):
            w[i] = BytesToWord(sMsg, blk * 64 + i * 4 + 1)
        for i in range(16, 80):
            w[i] = Rotl32(i32(w[i - 3] ^ w[i - 8] ^ w[i - 14] ^ w[i - 16]), 1)

        a, b, c, d, e = h0, h1, h2, h3, h4

        for i in range(80):
            if i < 20:
                f = i32((b & c) | ((~b) & d))
                kk = U32(1518500249.0)
            elif i < 40:
                f = i32(b ^ c ^ d)
                kk = U32(1859775393.0)
            elif i < 60:
                f = i32((b & c) | (b & d) | (c & d))
                kk = U32(2400959708.0)
            else:
                f = i32(b ^ c ^ d)
                kk = U32(3395469782.0)

            temp = Add32(Add32(Add32(Rotl32(a, 5), f), Add32(e, kk)), w[i])
            e = d
            d = c
            c = Rotl32(b, 30)
            b = a
            a = temp

        h0 = Add32(h0, a)
        h1 = Add32(h1, b)
        h2 = Add32(h2, c)
        h3 = Add32(h3, d)
        h4 = Add32(h4, e)

    return Hex32(h0) + Hex32(h1) + Hex32(h2) + Hex32(h3) + Hex32(h4)


if __name__ == "__main__":
    cases = ["", "a", "abc", "12345678", "Politecnica1",
             "The quick brown fox jumps over the lazy dog",
             "a" * 55, "a" * 56, "a" * 63, "a" * 64, "a" * 65, "a" * 1000,
             "Contrasena#2026 con acentos: aeiou"]
    ok = True
    for c in cases:
        mine = Sha1Hex(c)
        ref = hashlib.sha1(c.encode("utf-8")).hexdigest().upper()
        status = "OK " if mine == ref else "FAIL"
        if mine != ref:
            ok = False
        print(f"{status} len={len(c):5d} {mine} {ref}")
    print()
    print("RESULTADO:", "TODAS LAS PRUEBAS PASARON" if ok else "HAY FALLOS")
