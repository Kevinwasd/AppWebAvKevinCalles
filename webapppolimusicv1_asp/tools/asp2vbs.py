#!/usr/bin/env python3
"""
Convierte archivos .asp (VBScript) a .vbs para validarlos con cscript.
- Elimina el texto HTML que queda fuera de los bloques <% %>
- <%@ ... %>  -> se descarta (directiva)
- <% ... %>   -> codigo
- <%= expr %> -> Response.Write(expr)
- <!--#include ... --> -> se descarta (las dependencias se pasan aparte)

Se usa para detectar errores de compilacion (sintaxis y variables no
declaradas cuando el archivo usa Option Explicit).
"""
import re
import sys
import os

TAG = re.compile(r"<%(.*?)%>", re.S)


def asp_to_vbs(text: str) -> str:
    out = []
    for m in TAG.finditer(text):
        block = m.group(1)
        if block.startswith("@"):
            continue                      # directiva @LANGUAGE / @CODEPAGE
        if block.startswith("="):
            out.append("Call Response.Write(" + block[1:].strip().rstrip(";") + ")")
        else:
            out.append(block)
    return "\n".join(out)


def main() -> int:
    files = sys.argv[1:]
    rc = 0
    for f in files:
        if not os.path.exists(f):
            print("MISSING", f)
            rc = 1
            continue
        text = open(f, "r", encoding="utf-8", errors="replace").read()
        vbs = asp_to_vbs(text)
        target = os.path.join("/tmp/pmtest", os.path.basename(f) + ".vbs")
        open(target, "w", encoding="utf-8").write(vbs)
        print("OK ->", target, len(vbs), "chars")
    return rc


if __name__ == "__main__":
    sys.exit(main())
