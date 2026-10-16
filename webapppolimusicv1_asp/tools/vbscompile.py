"""
vbscompile.py - Prepara archivos .vbs a partir de las paginas ASP para
compilarlos con el motor real de VBScript (MSScriptControl).

Genera, para cada .asp, un .vbs con el MISMO numero de lineas que el
archivo original (las lineas de HTML quedan en blanco), de modo que los
errores de compilacion se puedan situar en la linea exacta.

Ademas, para cada pagina, genera el "cierre" completo (pagina + todos sus
#include resueltos) que es lo que IIS compila realmente.

Los objetos intrinsecos de ASP (Response, Request, Server, Session,
Application) se declaran para que 'Option Explicit' no los marque como
variables no declaradas.
"""
from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vbslint as V  # noqa: E402
import vbscallcheck as C  # noqa: E402

ROOT = V.ROOT
OUT = os.path.join(os.environ.get("TEMP", "/tmp"), "pmcompile")

INTRINSIC_DECL = "Dim Response, Request, Server, Session, Application, ObjectContext"

OPTION_RE = re.compile(r"^\s*Option\s+Explicit\s*$", re.I)


def build_code(path: str, mode: str = "page") -> list[str]:
    """Lineas de codigo listas para compilar.

    mode = "page"   -> PAGINA completa: cabecera con Option Explicit y la
                       declaracion de los objetos de ASP, e includes resueltos.
    mode = "single" -> ARCHIVO SUELTO: cabecera solo con los objetos de ASP
                       (y Option Explicit si el archivo lo declara), sin
                       includes. Sirve para localizar errores de sintaxis
                       dentro de un include.
    mode = "inline" -> fragmento para insertar dentro de una pagina: sin
                       cabecera, porque la aporta el archivo contenedor.
    """
    lines = V.asp_to_lines(path)
    declares_option = any(OPTION_RE.match(l) for l in lines)

    if mode == "page":
        text = open(path, "r", encoding="utf-8", errors="replace").read()
        extra: list[tuple[int, str]] = []
        for m in V.INCLUDE.finditer(text):
            inc = m.group(1).replace("\\", "/").lstrip("/")
            cand = os.path.join(os.path.dirname(path), inc)
            if not os.path.exists(cand):
                cand = os.path.join(ROOT, inc)
            if os.path.exists(cand):
                lineno = text.count("\n", 0, m.start())
                extra.append((lineno, cand))
        for lineno, cand in reversed(extra):
            sub = build_code(cand, "inline")
            lines[lineno] = "' @@BEGIN " + os.path.relpath(cand, ROOT) + " @@"
            lines.insert(lineno + 1, "\r\n".join(sub))
            lines.insert(lineno + 2, "' @@END@@")

    # Option Explicit debe ser la PRIMERA instruccion del script: se
    # neutraliza donde aparezca y se coloca al principio cuando corresponde.
    lines = ["" if OPTION_RE.match(line) else line for line in lines]

    if mode == "inline":
        return lines

    head: list[str] = []
    if mode == "page" or declares_option:
        head.append("Option Explicit")
    head.append(INTRINSIC_DECL)
    return head + lines


def write_vbs(target: str, lines: list[str]) -> None:
    """Escribe el .vbs con saltos CRLF reales (newline='' evita duplicarlos)."""
    with open(target, "w", encoding="utf-8", newline="") as fh:
        fh.write("\r\n".join(lines))


def main() -> None:
    os.makedirs(OUT, exist_ok=True)
    for f in os.listdir(OUT):
        os.remove(os.path.join(OUT, f))

    manifest = []

    # 1. Cada archivo por separado
    for base, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "tools", "assets",
                                                "uploads", "__pycache__")]
        for n in sorted(names):
            if not n.endswith(".asp"):
                continue
            p = os.path.join(base, n)
            rel = os.path.relpath(p, ROOT).replace("\\", "/")
            tag = rel.replace("/", "__").replace(".asp", "")
            target = os.path.join(OUT, "file_" + tag + ".vbs")
            write_vbs(target, build_code(p, "single"))
            manifest.append((target, "ARCHIVO " + rel))

    # 2. El cierre completo de cada pagina
    for base, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "tools", "assets",
                                                "uploads", "__pycache__")]
        for n in sorted(names):
            if not n.endswith(".asp"):
                continue
            p = os.path.join(base, n)
            rel = os.path.relpath(p, ROOT).replace("\\", "/")
            if rel.startswith("includes/"):
                continue
            tag = rel.replace("/", "__").replace(".asp", "")
            target = os.path.join(OUT, "page_" + tag + ".vbs")
            write_vbs(target, build_code(p, "page"))
            manifest.append((target, "PAGINA  " + rel))

    with open(os.path.join(OUT, "manifest.txt"), "w", encoding="utf-8") as fh:
        for target, label in manifest:
            fh.write(label + "\t" + target + "\n")

    print("Generados", len(manifest), "archivos en", OUT)


if __name__ == "__main__":
    main()
