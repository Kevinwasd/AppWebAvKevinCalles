"""
vbslint.py - Validador estatico de VBScript para ASP clasico.

No puede sustituir al compilador de IIS, pero detecta los dos tipos de
error que aparecen en tiempo de COMPILACION:

  1. Estructura de bloques desbalanceada
     (If/End If, For/Next, Do/Loop, Function/End Function, ...)
  2. Variables no declaradas cuando el archivo usa Option Explicit

Ademas marca construcciones que no existen en VBScript (!=, &&, ||, //, ++).

Uso:
    python tools/vbslint.py <archivo.asp> [mas archivos...]
    python tools/vbslint.py --all
"""
from __future__ import annotations

import os
import re
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

TAG = re.compile(r"<%(.*?)%>", re.S)
INCLUDE = re.compile(r"<!--\s*#include\s+(?:file|virtual)\s*=\s*\"([^\"]+)\"\s*-->", re.I)

INTRINSICS = {
    "request", "response", "server", "session", "application",
    "objectcontext", "err", "true", "false", "nothing", "null", "empty",
    "vbcrlf", "vbcr", "vblf", "vbtab", "vbnullchar", "vbnullstring",
    "vbobjecterror", "vbtab", "vbnewline", "vbfirstfourdays",
    "adopendynamic", "adopenkeyset", "adopenstatic", "adlockreadonly",
    "me",
}

VBS_BUILTINS = {
    "abs", "array", "asc", "ascb", "ascw", "atn", "cbool", "cbyte", "ccur",
    "cdate", "cdbl", "chr", "chrb", "chr", "cint", "clng", "cos", "createobject",
    "csng", "cstr", "date", "dateadd", "datediff", "datepart", "dateserial",
    "datevalue", "day", "eval", "exp", "filter", "fix", "formatcurrency",
    "formatdatetime", "formatnumber", "formatpercent", "getobject", "getlocale",
    "hex", "hour", "inputbox", "instr", "instrrev", "int", "isarray", "isdate",
    "isempty", "isnull", "isnumeric", "isobject", "join", "lbound", "lcase",
    "left", "len", "lenb", "loadpicture", "log", "ltrim", "mid", "midb",
    "minute", "month", "monthname", "msgbox", "now", "oct", "replace", "rgb",
    "right", "rnd", "round", "rtrim", "second", "setlocale", "sgn", "sin",
    "space", "split", "sqr", "strcomp", "string", "strreverse", "tan", "time",
    "timer", "timeserial", "timevalue", "trim", "typename", "ubound", "ucase",
    "vartype", "weekday", "weekdayname", "year", "clng", "cint", "csng",
}

BLOCK_OPEN = [
    (re.compile(r"^\s*if\b.*\bthen\s*$", re.I), "if"),
    (re.compile(r"^\s*for\b", re.I), "for"),
    (re.compile(r"^\s*do\b", re.I), "do"),
    (re.compile(r"^\s*select\s+case\b", re.I), "select"),
    (re.compile(r"^\s*with\b", re.I), "with"),
    (re.compile(r"^\s*(public\s+|private\s+)?function\b", re.I), "proc"),
    (re.compile(r"^\s*(public\s+|private\s+)?sub\b", re.I), "proc"),
    (re.compile(r"^\s*(public\s+|private\s+)?property\b", re.I), "proc"),
    (re.compile(r"^\s*class\b", re.I), "class"),
]

BLOCK_CLOSE = [
    (re.compile(r"^\s*end\s+if\b", re.I), "if"),
    (re.compile(r"^\s*next\b", re.I), "for"),
    (re.compile(r"^\s*loop\b", re.I), "do"),
    (re.compile(r"^\s*end\s+select\b", re.I), "select"),
    (re.compile(r"^\s*end\s+with\b", re.I), "with"),
    (re.compile(r"^\s*end\s+(function|sub|property)\b", re.I), "proc"),
    (re.compile(r"^\s*end\s+class\b", re.I), "class"),
]

BAD_PATTERNS = [
    (re.compile(r"[^<>!=]!=[^=]"), "use '<>' en VBScript, no '!='"),
    (re.compile(r"&&"), "use 'And' en VBScript, no '&&'"),
    (re.compile(r"\|\|"), "use 'Or' en VBScript, no '||'"),
    (re.compile(r"//(?!/)"), "use ''' para comentarios, no '//'"),
    (re.compile(r"\+\+|\-\-"), "VBScript no tiene '++' / '--'"),
    (re.compile(r"\bvar\s+\w+\s*="), "VBScript no tiene 'var'"),
]


def strip_comment(line: str) -> str:
    """Elimina el comentario ' fuera de las cadenas."""
    out = []
    in_str = False
    i = 0
    while i < len(line):
        ch = line[i]
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        elif ch == "'" and not in_str:
            break
        else:
            out.append(ch)
        i += 1
    return "".join(out)


def mask_strings(line: str) -> str:
    """Sustituye el contenido de las cadenas por X para no confundir palabras."""
    out = []
    in_str = False
    for ch in line:
        if ch == '"':
            in_str = not in_str
            out.append(ch)
        else:
            out.append("X" if in_str else ch)
    return "".join(out)


def split_statements(line: str) -> list[str]:
    """Divide por ':' respetando las cadenas."""
    parts, cur, in_str = [], [], False
    for ch in line:
        if ch == '"':
            in_str = not in_str
            cur.append(ch)
        elif ch == ":" and not in_str:
            parts.append("".join(cur))
            cur = []
        else:
            cur.append(ch)
    parts.append("".join(cur))
    return [p for p in parts if p.strip()]


def asp_to_lines(path: str) -> list[str]:
    """Devuelve las lineas de codigo VBScript del .asp.

    El texto HTML entre bloques se reemplaza por lineas vacias para que la
    numeracion de lineas coincida con la del archivo original.
    """
    text = open(path, "r", encoding="utf-8", errors="replace").read()
    result = [""] * (text.count("\n") + 1)

    for m in TAG.finditer(text):
        code = m.group(1)
        if code.startswith("@"):
            continue                      # directiva @LANGUAGE / @CODEPAGE
        if code.startswith("="):
            code = "Call Response.Write(" + code[1:].strip().rstrip(";") + ")"

        first = text.count("\n", 0, m.start())          # linea 0-based
        for i, chunk in enumerate(code.split("\n")):
            idx = first + i
            chunk = chunk.replace("\r", "").rstrip()
            if idx >= len(result):
                result.extend([""] * (idx - len(result) + 1))
            if chunk.strip():
                result[idx] = (result[idx] + "\n" + chunk) if result[idx] else chunk
    return result


def resolve_includes(path: str, seen: set | None = None) -> list[str]:
    """Rutas (absolutas) de los archivos incluidos por #include."""
    seen = seen or set()
    real = os.path.abspath(path)
    if real in seen:
        return []
    seen.add(real)
    found = []
    text = open(path, "r", encoding="utf-8", errors="replace").read()
    for m in INCLUDE.finditer(text):
        inc = m.group(1).replace("\\", "/").lstrip("/")
        cand = os.path.join(os.path.dirname(path), inc)
        if not os.path.exists(cand):
            cand = os.path.join(ROOT, inc)
        if os.path.exists(cand):
            found.append(cand)
            found.extend(resolve_includes(cand, seen))
    return found


def harvest_declarations(path: str) -> set[str]:
    """Nombres declarados en un archivo, sus includes y la libreria comun."""
    names: set[str] = set()
    files = [path] + resolve_includes(path)

    # Toda pagina incluye la libreria comun de includes/: sus Dim y Const
    # son globales para el resto del codigo.
    inc_dir = os.path.join(ROOT, "includes")
    if os.path.isdir(inc_dir):
        for base, _dirs, fnames in os.walk(inc_dir):
            for n in fnames:
                if n.lower().endswith(".asp"):
                    files.append(os.path.join(base, n))

    for f in files:
        for _lineno, line in code_lines(f):
            masked = mask_strings(line)
            for stmt in split_statements(masked):
                s = stmt.strip()
                low = s.lower()
                m = re.match(r"^\s*(?:dim|public|private|static)\s+(.+)$", low)
                if m:
                    body = re.sub(r"^\s*(?:dim|public|private|static)\s+", "", low)
                    for piece in body.split(","):
                        nm = re.match(r"^\s*([a-z_]\w*)", piece.strip())
                        if nm:
                            names.add(nm.group(1))
                    continue
                m = re.match(r"^\s*const\s+([a-z_]\w*)", low)
                if m:
                    names.add(m.group(1))
                    continue
                m = re.match(r"^\s*(?:public\s+|private\s+)?(?:function|sub|property)\s+([a-z_]\w*)", low)
                if m:
                    names.add(m.group(1))
                    params = re.search(r"\((.*)\)", s)
                    if params:
                        for p in params.group(1).split(","):
                            p = re.sub(r"^(byval|byref|optional)\s+", "", p.strip(), flags=re.I)
                            p = re.sub(r"\s*=.*$", "", p).strip()
                            if re.match(r"^[A-Za-z_]\w*$", p):
                                names.add(p.lower())
                    continue
                m = re.match(r"^\s*for\s+(?:each\s+)?([a-z_]\w*)", low)
                if m:
                    names.add(m.group(1))
    return names


def code_lines(path: str) -> list[tuple[int, str]]:
    """Lineas de codigo (sin comentarios, sin continuaciones) numeradas."""
    raw = asp_to_lines(path)
    flat: list[tuple[int, str]] = []
    for idx, line in enumerate(raw, start=1):
        for sub in line.split("\n"):
            flat.append((idx, sub))

    out: list[tuple[int, str]] = []
    pending = ""
    start = 0
    for idx, line in flat:
        s = line.rstrip()
        if not pending:
            start = idx
        stripped = strip_comment(s).rstrip()
        if stripped.endswith(" _") or stripped.endswith("\t_"):
            pending += stripped[:-1]
            continue
        full = pending + stripped
        pending = ""
        if full.strip():
            out.append((start, full))
    if pending.strip():
        out.append((start, pending))
    return out


def comment_hides_code(line: str) -> str | None:
    """Detecta un comentario que oculta codigo en la misma linea.

    Suele indicar que dos lineas se unieron por error (por ejemplo al
    editar el archivo): ' comentario ... Function Foo()  ->  el compilador
    descarta la declaracion y el bloque queda descuadrado.
    """
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            tail = line[i + 1:].strip()
            if re.search(r"\b(end\s+)?(function|sub|property|class)\b", tail, re.I) \
               or re.match(r"^(dim|const|if\b.*\bthen|for\b|do\b|select\s+case|with\b|next\b|loop\b)",
                           tail, re.I):
                return tail[:70]
            return None
    return None


def lint(path: str) -> list[str]:
    errors: list[str] = []
    rel = os.path.relpath(path, ROOT).replace("\\", "/")
    declared: set[str] = set(INTRINSICS) | VBS_BUILTINS
    declared |= harvest_declarations(path)
    stack: list[tuple[str, int, str]] = []

    lines = code_lines(path)

    for lineno, line in lines:
        masked = mask_strings(line)
        low = masked.strip().lower()

        for pat, msg in BAD_PATTERNS:
            if pat.search(masked):
                errors.append(f"{rel}:{lineno}: {msg}  ->  {line.strip()[:90]}")

        hidden = comment_hides_code(line)
        if hidden:
            errors.append(
                f"{rel}:{lineno}: comentario que oculta codigo ('{hidden}') "
                f"-> probable union accidental de dos lineas"
            )

        for stmt in split_statements(masked):
            s = stmt.strip()
            if not s:
                continue
            slow = s.lower()

            # ---- cierres de bloque -------------------------------------
            closed = False
            for pat, kind in BLOCK_CLOSE:
                if pat.match(slow):
                    if not stack:
                        errors.append(f"{rel}:{lineno}: cierre '{s}' sin bloque abierto")
                    else:
                        k, sl, txt = stack.pop()
                        if k != kind:
                            errors.append(
                                f"{rel}:{lineno}: '{s}' cierra un bloque '{k}' "
                                f"abierto en la linea {sl} ('{txt}')"
                            )
                    closed = True
                    break
            if closed:
                continue

            # ---- aperturas de bloque -----------------------------------
            if re.match(r"^\s*else\s*$", slow) or re.match(r"^\s*elseif\b", slow) \
               or re.match(r"^\s*case\b", slow):
                continue

            opened = False
            for pat, kind in BLOCK_OPEN:
                if pat.match(slow):
                    # If de una sola linea: "If x Then sentencia"
                    if kind == "if":
                        m = re.match(r"^\s*if\b(.*)\bthen\b(.*)$", slow)
                        if m and m.group(2).strip():
                            opened = True   # no genera bloque
                            break
                    stack.append((kind, lineno, s[:60]))
                    opened = True
                    break
            if opened:
                # aun puede contener una declaracion implicita (For i = ...)
                m = re.match(r"^\s*for\s+(?:each\s+)?([A-Za-z_]\w*)", slow)
                if m:
                    declared.add(m.group(1))
                m = re.match(r"^\s*(?:public\s+|private\s+)?(?:function|sub|property)\s+([A-Za-z_]\w*)", slow)
                if m:
                    declared.add(m.group(1))
                    params = re.search(r"\((.*)\)", s)
                    if params:
                        for p in params.group(1).split(","):
                            p = p.strip()
                            p = re.sub(r"^(byval|byref|optional)\s+", "", p, flags=re.I)
                            p = re.sub(r"\s*=.*$", "", p).strip()
                            if re.match(r"^[A-Za-z_]\w*$", p):
                                declared.add(p.lower())
                continue

            # ---- declaraciones ------------------------------------------
            m = re.match(r"^\s*(?:public\s+|private\s+|dim\s+|static\s+)(.+)$", slow)
            if m and (slow.startswith("dim") or slow.startswith("public")
                      or slow.startswith("private") or slow.startswith("static")):
                body = re.sub(r"^\s*(?:public|private|dim|static)\s+", "", slow)
                for piece in body.split(","):
                    nm = re.match(r"^\s*([A-Za-z_]\w*)", piece.strip())
                    if nm:
                        declared.add(nm.group(1))
                continue

            m = re.match(r"^\s*const\s+([A-Za-z_]\w*)", slow)
            if m:
                declared.add(m.group(1))
                continue

            m = re.match(r"^\s*option\s+explicit", slow)
            if m:
                continue

            # ---- asignaciones -------------------------------------------
            m = re.match(r"^\s*(?:set\s+)?([A-Za-z_]\w*)\s*(\([^)]*\))?\s*=", s)
            if m and not re.match(r"^\s*(if|elseif|for|while|do|select|case|with|function|sub|class|end)\b", slow):
                name = m.group(1).lower()
                if name not in declared:
                    errors.append(
                        f"{rel}:{lineno}: variable no declarada '{m.group(1)}' "
                        f"(Option Explicit)  ->  {s[:80]}"
                    )
                continue

    for kind, lineno, txt in stack:
        errors.append(f"{rel}:{lineno}: bloque '{kind}' sin cerrar  ->  {txt}")

    return errors


def main() -> int:
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        return 1
    if args == ["--all"]:
        files = []
        for base, _dirs, names in os.walk(ROOT):
            if ".git" in base or "tools" in base or "assets" in base:
                continue
            for n in names:
                if n.lower().endswith(".asp"):
                    files.append(os.path.join(base, n))
        files.sort()
    else:
        files = args

    total = 0
    for f in files:
        errs = lint(f)
        rel = os.path.relpath(f, ROOT).replace("\\", "/")
        if errs:
            total += len(errs)
            print(f"\n### {rel}  ({len(errs)} problema(s))")
            for e in errs:
                print("   ", e)
        else:
            print(f"OK   {rel}")
    print()
    print(f"TOTAL: {total} problema(s) en {len(files)} archivo(s)")
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main())
