"""
Reorganiza las paginas de admin/ y api/ a la raiz de la aplicacion.

Motivo: en ASP clasico, #include file="../..." requiere la directiva
enableParentPaths (deshabilitada por omision en IIS y solo configurable
por el administrador). Manteniendo todas las paginas en la raiz de la
aplicacion, las inclusiones son siempre "includes/..." y la aplicacion
funciona en cualquier servidor IIS sin configuracion adicional.
"""
import os
import re
import shutil

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

MOVES = {
    "admin/index.asp": "admin.asp",
    "admin/songs.asp": "admin_songs.asp",
    "admin/artists.asp": "admin_artists.asp",
    "admin/albums.asp": "admin_albums.asp",
    "admin/playlists.asp": "admin_playlists.asp",
    "admin/users.asp": "admin_users.asp",
    "api/view.asp": "api_view.asp",
    "api/play.asp": "api_play.asp",
    "api/playlist_add.asp": "api_playlist_add.asp",
    "api/playlist_remove.asp": "api_playlist_remove.asp",
    "api/playlist_pick.asp": "api_playlist_pick.asp",
}

# Reemplazos globales de texto (rutas y enlaces)
REPLACEMENTS = [
    (r'<!--#include file="\.\./includes/', '<!--#include file="includes/'),
    (r'U\("admin/index\.asp"\)', 'U("admin.asp")'),
    (r'U\("admin/songs\.asp', 'U("admin_songs.asp'),
    (r'U\("admin/artists\.asp', 'U("admin_artists.asp'),
    (r'U\("admin/albums\.asp', 'U("admin_albums.asp'),
    (r'U\("admin/playlists\.asp', 'U("admin_playlists.asp'),
    (r'U\("admin/users\.asp', 'U("admin_users.asp'),
    (r'"admin/index\.asp"', '"admin.asp"'),
    (r'"admin/songs\.asp"', '"admin_songs.asp"'),
    (r'"admin/artists\.asp"', '"admin_artists.asp"'),
    (r'"admin/albums\.asp"', '"admin_albums.asp"'),
    (r'"admin/playlists\.asp"', '"admin_playlists.asp"'),
    (r'"admin/users\.asp"', '"admin_users.asp"'),
    (r"BASE \+ 'api/view\.asp'", "BASE + 'api_view.asp'"),
    (r"BASE \+ 'api/play\.asp'", "BASE + 'api_play.asp'"),
    (r"BASE \+ 'api/playlist_add\.asp'", "BASE + 'api_playlist_add.asp'"),
    (r"BASE \+ 'api/playlist_remove\.asp'", "BASE + 'api_playlist_remove.asp'"),
    (r"BASE \+ 'api/playlist_pick\.asp'", "BASE + 'api_playlist_pick.asp'"),
]


def main() -> None:
    # 1. Mover los archivos
    for src, dst in MOVES.items():
        sp = os.path.join(ROOT, src)
        dp = os.path.join(ROOT, dst)
        if not os.path.exists(sp):
            print("omitido (no existe):", src)
            continue
        if os.path.exists(dp):
            os.remove(dp)
        shutil.move(sp, dp)
        print("movido:", src, "->", dst)

    # 2. Aplicar los reemplazos en todos los archivos de la aplicacion
    exts = (".asp", ".js", ".css", ".md", ".ps1", ".xml")
    changed = 0
    for base, dirs, names in os.walk(ROOT):
        dirs[:] = [d for d in dirs if d not in (".git", "__pycache__", "assets", "uploads")]
        for n in names:
            if not n.endswith(exts):
                continue
            p = os.path.join(base, n)
            try:
                text = open(p, "r", encoding="utf-8").read()
            except UnicodeDecodeError:
                continue
            original = text
            for pat, rep in REPLACEMENTS:
                text = re.sub(pat, rep, text)
            if text != original:
                open(p, "w", encoding="utf-8").write(text)
                changed += 1
                print("actualizado:", os.path.relpath(p, ROOT))

    # 3. Eliminar las carpetas vacias
    for d in ("admin", "api"):
        p = os.path.join(ROOT, d)
        if os.path.isdir(p):
            rest = os.listdir(p)
            if not rest:
                os.rmdir(p)
                print("carpeta eliminada:", d)
            else:
                print("carpeta NO vacia, se conserva:", d, rest)

    print("\nArchivos actualizados:", changed)


if __name__ == "__main__":
    main()
