"""
vbscallcheck.py - Verificacion de referencias entre archivos ASP.

IIS compila cada pagina junto con TODOS sus #include en un unico script.
Este verificador reproduce ese cierre de inclusiones y comprueba que:

  1. Cada #include file="..." apunta a un archivo existente.
  2. Cada funcion / sub invocada esta definida en alguna parte del script
     compilado (pagina + includes).
  3. Cada Sub invocada como sentencia tambien existe.

Es el equivalente mas cercano a un "enlazador" sin poder ejecutar IIS.

Uso:  python tools/vbscallcheck.py
"""
from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vbslint as V  # noqa: E402

ROOT = V.ROOT

KEYWORDS = {
    "dim", "const", "set", "if", "then", "else", "elseif", "end", "for", "each",
    "next", "do", "while", "until", "loop", "select", "case", "with", "function",
    "sub", "property", "class", "exit", "option", "explicit", "call", "redim",
    "erase", "on", "error", "resume", "goto", "rem", "public", "private",
    "static", "new", "byval", "byref", "let", "stop", "wend", "and", "or",
    "not", "xor", "mod", "is", "to", "step",
}

CALL_RE = re.compile(r"(?<![\w.])([A-Za-z_]\w*)\s*\(")
STMT_RE = re.compile(r"^\s*([A-Za-z_]\w*)(.*)$", re.S)

PROC_RE = re.compile(
    r"^\s*(?:public\s+|private\s+)?(?:function|sub|property\s+(?:get|let|set)\s+)([A-Za-z_]\w*)",
    re.I,
)


def closure(path: str) -> list[str]:
    """La pagina mas todos los archivos que incluye (recursivo)."""
    out = []
    seen = set()

    def walk(p: str) -> None:
        real = os.path.abspath(p)
        if real in seen:
            return
        seen.add(real)
        out.append(real)
        for inc in V.resolve_includes(p):
            walk(inc)

    walk(path)
    return out


def check_page(path: str) -> list[str]:
    problems: list[str] = []
    files = closure(path)
    rel = os.path.relpath(path, ROOT).replace("\\", "/")

    defined: set[str] = set()
    for f in files:
        for _ln, line in V.code_lines(f):
            m = PROC_RE.match(V.mask_strings(line))
            if m:
                defined.add(m.group(1).lower())

    declared = set(V.INTRINSICS) | V.VBS_BUILTINS | set(KEYWORDS)
    for f in files:
        declared |= V.harvest_declarations(f)

    for f in files:
        frel = os.path.relpath(f, ROOT).replace("\\", "/")
        for lineno, line in V.code_lines(f):
            masked = V.mask_strings(line)
            for stmt in V.split_statements(masked):
                s = stmt.strip()
                if not s:
                    continue

                # --- llamadas tipo funcion: Nombre(...) ---
                for m in CALL_RE.finditer(s):
                    name = m.group(1).lower()
                    if name in declared or name in defined:
                        continue
                    if name in KEYWORDS:
                        continue
                    problems.append(
                        f"{frel}:{lineno}: funcion no definida '{m.group(1)}()' "
                        f"(pagina {rel})  ->  {s[:70]}"
                    )

                # --- llamadas tipo sentencia: Nombre arg1, arg2 ---
                m = STMT_RE.match(s)
                if m:
                    name = m.group(1)
                    rest = m.group(2)
                    # '.algo' es una propiedad/metodo y '(algo)' una llamada:
                    # en ambos casos ya se analizan por otra via.
                    if rest[:1] in (".", "("):
                        continue
                    low = name.lower()
                    if low in KEYWORDS or low in declared or low in defined:
                        continue
                    if "=" in s:
                        continue
                    problems.append(
                        f"{frel}:{lineno}: sub no definida '{name}' "
                        f"(pagina {rel})  ->  {s[:70]}"
                    )
    return problems


def main() -> int:
    pages = []
    for base, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "tools", "assets", "uploads")]
        for n in names:
            if n.lower().endswith(".asp"):
                pages.append(os.path.join(base, n))
    pages.sort()

    total = 0
    for p in pages:
        errs = sorted(set(check_page(p)))
        if errs:
            total += len(errs)
            print(f"\n### {os.path.relpath(p, ROOT)}  ({len(errs)})")
            for e in errs:
                print("   ", e)
    print(f"\nTOTAL: {total} referencia(s) sin resolver en {len(pages)} paginas")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
