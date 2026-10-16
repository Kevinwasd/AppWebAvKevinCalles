"""
htmlcheck.py - Comprobacion del equilibrio de etiquetas HTML.

Extrae el marcado literal (fuera de los bloques <% %>) de cada pagina y
verifica que las etiquetas de apertura y cierre esten balanceadas.

Los bloques condicionales (<% If %> ... <% End If %>) se cuentan como
ramas independientes del mismo nivel, por lo que se tolera un pequeno
desajuste: lo importante es detectar etiquetas claramente desbalanceadas
(por ejemplo un <div> sin cerrar).
"""
from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vbslint as V  # noqa: E402

ROOT = V.ROOT

VOID = {"area", "base", "br", "col", "embed", "hr", "img", "input", "link",
        "meta", "param", "source", "track", "wbr"}

TAG_RE = re.compile(r"<(/?)([a-zA-Z][a-zA-Z0-9]*)([^>]*?)(/?)>", re.S)


def literal_markup(path: str) -> str:
    text = open(path, "r", encoding="utf-8", errors="replace").read()
    out = []
    pos = 0
    for m in V.TAG.finditer(text):
        out.append(text[pos:m.start()])
        pos = m.end()
    out.append(text[pos:])
    return "".join(out)


def check(path: str) -> tuple[int, list[str]]:
    markup = literal_markup(path)
    stack: list[tuple[str, int]] = []
    problems: list[str] = []
    for m in TAG_RE.finditer(markup):
        closing, name, attrs, selfclose = m.group(1), m.group(2).lower(), m.group(3), m.group(4)
        if name in VOID or selfclose == "/":
            continue
        line = markup.count("\n", 0, m.start()) + 1
        if closing:
            if stack and stack[-1][0] == name:
                stack.pop()
            elif any(t == name for t, _ in stack):
                # Cierra una etiqueta anterior: se descartan las intermedias
                while stack and stack[-1][0] != name:
                    t, l = stack.pop()
                    problems.append(f"linea {l}: <{t}> sin cerrar antes de </{name}>")
                if stack:
                    stack.pop()
            else:
                problems.append(f"linea {line}: </{name}> sin apertura")
        else:
            stack.append((name, line))

    for t, l in stack:
        if t in ("html", "body", "head"):
            continue
        problems.append(f"linea {l}: <{t}> sin cerrar")
    return len(stack), problems


def main() -> int:
    # layout_top.asp abre el marco (div.app, div.main, main) y layout_bottom.asp
    # lo cierra: son un par y no se analizan por separado.
    LAYOUT_PAIR = {"layout_top.asp", "layout_bottom.asp"}

    pages = []
    for base, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "tools", "assets", "uploads")]
        for n in names:
            if n.lower().endswith(".asp"):
                pages.append(os.path.join(base, n))
    pages.sort()

    total = 0
    for p in pages:
        rel = os.path.relpath(p, ROOT).replace("\\", "/")
        if os.path.basename(p) in LAYOUT_PAIR:
            print(f"SKIP {rel}  (mitad del marco comun: layout_top + layout_bottom)")
            continue
        _, problems = check(p)
        if problems:
            total += len(problems)
            print(f"\n### {rel}  ({len(problems)})")
            for x in problems[:12]:
                print("   ", x)
        else:
            print(f"OK   {rel}")
    print(f"\nTOTAL: {total} problema(s) de marcado")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
