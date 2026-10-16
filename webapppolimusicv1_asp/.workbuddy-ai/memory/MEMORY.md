# PoliMusic — Notas permanentes del proyecto

## Stack y convenciones
- **ASP Clásico (VBScript)** + **SQL Server** + **IIS**. Sin frameworks ni
  componentes COM externos (solo ADO y Scripting.FileSystemObject).
- Carpeta de despliegue acordada: `C:\inetpub\wwwroot\webapppolimusicv1_asp`.
- Color principal: **turquesa** `#14b8a6` (variables CSS `--pm-*` en
  `assets/css/style.css`). Interfaz oscura tipo Spotify.
- Base de datos `BDD_PoliMusic`; cadena de conexión solicitada:
  `Server=localhost; Database=BDD_PoliMusic; User Id=usr_polimusic;
  Password=Politecnica1; Encrypt=true; TrustServerCertificate=true`
  (traducida a OLE DB en `includes/config.asp` + `DbConnectionString()`).

## Reglas obligatorias del código
1. **Todas las páginas ASP van en la raíz de la aplicación.** Nunca usar
   `#include file="../..."`: requiere `enableParentPaths`, deshabilitado por
   omisión en IIS. Las inclusiones son siempre `includes/...`.
2. **Archivos fuente en ASCII puro.** Acentos como entidades HTML (`&oacute;`)
   y en JavaScript como `\u00bf`. Nunca guardar bytes > 127 en `.asp/.js/.css`.
3. **No declarar la sección `<asp>` en `web.config`** (riesgo de error 500.19).
   Los ajustes de ASP se aplican con `appcmd` desde `deploy.ps1`.
4. `Option Explicit` en todos los archivos: declarar siempre las variables.
5. Toda entrada de usuario pasa por `SqlStr()` / `SqlInt()` antes de ir a SQL, y
   toda salida por `H()`.
6. `Param()` consulta la QueryString antes que `Request.Form`; en peticiones
   multipart usar siempre `UFField()`.
7. Los formularios que combinan guardar/eliminar usan un campo aparte
   (`do_delete=1`) en vez de dos valores para `action`.

## Estructura
- `includes/` — librería compartida (config, db, util, sha1, auth, upload, ui,
  icons, actions, playlists_logic, layout_top/bottom, views/*).
- `assets/` — css, js, y carpetas de recursos (`SongFiles`, `UserImages`,
  `PlaylistImages`, `AlbumImages`, `ArtistImages`).
- `tools/` — verificadores (no se despliegan): `vbslint.py`, `vbscallcheck.py`,
  `htmlcheck.py`, `sha1_check.py`, `sql_validation.sql`, `seed_usuarios.sql`.

## Ejecutar las validaciones antes de dar algo por terminado
```
python tools/vbslint.py --all      # bloques y variables
python tools/vbscallcheck.py       # includes y funciones
python tools/htmlcheck.py          # etiquetas HTML
python tools/sha1_check.py         # SHA-1 vs hashlib
node --check assets/js/app.js

# Compilacion con el MOTOR REAL de VBScript (imprescindible):
python tools/vbscompile.py
& "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" `
      -NoProfile -ExecutionPolicy Bypass -File tools\vbscheck.ps1
```

## Trampas de VBScript ya encontradas (no repetir)
1. **Identificadores que en minúsculas son palabras reservadas.** `sTop` es
   inválido porque `stop` es reservado. Se comprobaron los 196 identificadores
   declarados; usar `tools/vbscheck.ps1` tras cualquier cambio.
2. **Sufijo de tipo de VBA (`#`, `!`, `@`) en literales.** `4294967296#` compila
   en una asignación simple pero **rompe una condición `If`**. Usar `.0`.
3. **`Option Explicit` debe ser la primera instrucción del script compilado.**
   Por eso `config.asp` se incluye siempre el primero.
4. **Los `.vbs`/`.asp` de prueba deben escribirse con `newline=""` y unirse con
   `\r\n`**; en modo texto se genera `\r\r\n` y se rompen las continuaciones
   de línea (`_`), produciendo errores falsos.

## Herramientas de validación
`tools/vbscompile.py` traduce cada `.asp` a `.vbs` (página completa con sus
`#include` resueltos + archivo suelto) y `tools/vbscheck.ps1` lo compila con
`MSScriptControl` en PowerShell de 32 bits (`SysWOW64`), el mismo motor que IIS.
Es la única forma de detectar errores de compilación sin desplegar.

## Entorno de la máquina
- IIS activo y ASP clásico instalado; SQL Server en `localhost`.
- El usuario de trabajo **no es administrador**: no puede escribir en
  `C:\inetpub\wwwroot` ni ejecutar `Start-Process -Verb RunAs`. El despliegue
  requiere que la persona ejecute `desplegar.bat` / `deploy.ps1` elevado.
- `cscript.exe` está bloqueado por el sandbox: no se puede compilar VBScript
  localmente; usar los verificadores de `tools/`.
