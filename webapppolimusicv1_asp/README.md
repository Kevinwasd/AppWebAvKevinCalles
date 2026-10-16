# PoliMusic

Aplicación web de música tipo Spotify desarrollada en **ASP Clásico (VBScript)** con  
**SQL Server** como base de datos y **IIS** como servidor web.

Color principal: **turquesa** (`#14b8a6`).

---

## 1. Características

### Autenticación y roles

- **Inicio de sesión** y **registro de usuarios** con validaciones.
- Contraseñas almacenadas con **SHA-1** (implementación propia en VBScript, sin    
  componentes externos). El inicio de sesión acepta tanto valores con hash como los    
  valores en texto plano que trae el script `BDD_PoliMusic.sql`.
- Dos roles manejados con `TBL_USER.USER_TYPE` (catálogo `TBL_CATALOG`):
  - **Administrador** (`1`) — `userAdmin`
  - **Usuario normal** (`0`) — `epn.fis`, `userNormal`

| Rol            | Permisos                                                                                          |
| -------------- | ------------------------------------------------------------------------------------------------- |
| Administrador  | CRUD de **canciones, usuarios, artistas, álbumes y playlists**; acceso al panel de administración |
| Usuario normal | Escuchar canciones, consultar artistas y álbumes, y **CRUD de sus propias playlists**             |

### Reproductor de audio

- Barra inferior fija tipo Spotify: carátula, título, artista.
- Controles: reproducir/pausar, anterior, siguiente, **aleatorio**, **repetición**    
  (desactivada / todas / una), barra de progreso con búsqueda, volumen y silencio.
- **Cola de reproducción** construida a partir de la lista visible.
- **Contador de reproducciones** (`TBL_SONG.PLAYS`) actualizado en segundo plano.
- **Atajos de teclado**: `Espacio` reproducir/pausar, `Shift + →` siguiente,    
  `Shift + ←` anterior.
- Integración con la API **Media Session** (controles del sistema operativo).
- La navegación entre secciones se hace sin recargar la página, de modo que    
  **la música no se interrumpe**.

### Interfaz

- Tema oscuro tipo Spotify con acentos turquesa, totalmente responsive.
- Barra lateral con secciones de exploración y de administración (según el rol).
- Buscador global (canciones, artistas y álbumes) con resultados en vivo.
- Filtro local en la lista de canciones y ordenamiento (recientes / A-Z / más escuchadas).
- Imágenes faltantes reemplazadas por **marcadores SVG generados dinámicamente**    
  (`placeholder.asp`), por lo que nunca se muestra una imagen rota.

---

## 2. Requisitos

| Componente        | Detalle                                               |
| ----------------- | ----------------------------------------------------- |
| Sistema operativo | Windows 10 / 11 / Server                              |
| Servidor web      | IIS con la característica **ASP clásico** (`IIS-ASP`) |
| Base de datos     | SQL Server (instancia `localhost`)                    |
| Base de datos     | `BDD_PoliMusic` creada con `BDD_PoliMusic.sql`        |
| Navegador         | Chrome, Edge o Firefox actualizado                    |

---

## 3. Estructura del proyecto

```
webapppolimusicv1_asp/
├── default.asp              Punto de entrada (redirige a login o home)
├── login.asp                Inicio de sesión
├── register.asp             Registro de usuarios
├── logout.asp               Cierre de sesión
├── home.asp                 Página principal (shell de la aplicación)
├── profile.asp              Mi perfil
├── diagnostics.asp          Diagnóstico del entorno
├── placeholder.asp          Generador de imágenes SVG de respaldo
├── admin.asp                Panel de administración
├── admin_songs.asp          CRUD de canciones (+ artistas y álbumes)
├── admin_artists.asp        CRUD de artistas
├── admin_albums.asp         CRUD de álbumes (+ lista de canciones)
├── admin_playlists.asp      CRUD de playlists del sistema
├── admin_users.asp          CRUD de usuarios y roles
├── api_view.asp             Vistas parciales (navegación sin recarga)
├── api_play.asp             Registro de reproducciones (JSON)
├── api_playlist_add.asp     Agregar canción a playlist (JSON)
├── api_playlist_remove.asp  Quitar canción de playlist (JSON)
├── api_playlist_pick.asp    Selector de playlist (HTML)
├── includes/
│   ├── config.asp           Configuración y constantes
│   ├── db.asp               Capa de acceso a datos (ADO / OLE DB)
│   ├── util.asp             Utilidades, rutas y resolución de imágenes
│   ├── sha1.asp             SHA-1 en VBScript puro
│   ├── auth.asp             Sesión y control de acceso por roles
│   ├── upload.asp           Carga de archivos multipart sin componentes
│   ├── ui.asp               Componentes de interfaz reutilizables
│   ├── icons.asp            Iconografía SVG en línea
│   ├── actions.asp          Acciones POST (patrón POST/Redirect/GET)
│   ├── playlists_logic.asp  Reglas de negocio de las playlists
│   ├── layout_top.asp       Cabecera, barra lateral y barra superior
│   ├── layout_bottom.asp    Reproductor, modal y scripts
│   └── views/               Vistas: home, songs, albums, artists, playlists, search
├── assets/
│   ├── css/style.css        Hoja de estilos (tema turquesa)
│   ├── js/player.js         Motor del reproductor
│   ├── js/app.js            Navegación, modales y notificaciones
│   ├── img/favicon.svg
│   ├── SongFiles/           Archivos MP3
│   ├── UserImages/          Fotos de perfil
│   ├── PlaylistImages/      Portadas de playlists
│   ├── AlbumImages/         Portadas de álbumes
│   └── ArtistImages/        Fotos de artistas
├── uploads/                 Archivos subidos desde el panel
├── tools/                   Herramientas de validación (no se despliegan)
├── web.config               Configuración de IIS
└── deploy.ps1               Script de despliegue
```

> **Nota de diseño:** todas las páginas ASP viven en la raíz de la aplicación.>   
> En ASP clásico `#include file="../..."` requiere la directiva>   
> `enableParentPaths`, deshabilitada por omisión en IIS. Manteniendo las páginas>   
> en la raíz, las inclusiones son siempre `includes/...` y la aplicación funciona>   
> en cualquier servidor IIS **sin configuración adicional**.

---

## 4. Instalación

### 4.1 Base de datos

Ejecute el script provisto en la carpeta de trabajo (crea el login, la base, las  
tablas, los datos de ejemplo y las relaciones):

```sql
-- En SQL Server Management Studio o sqlcmd
:r BDD_PoliMusic.sql
```

Cree las cuentas de trabajo indicadas en los requerimientos:

```powershell
sqlcmd -S localhost -U usr_polimusic -P Politecnica1 -C -d BDD_PoliMusic -i tools\seed_usuarios.sql
```

### 4.2 Aplicación en IIS

Abra **PowerShell como Administrador** y ejecute:

```powershell
cd C:\Users\WebAvanzadas\Documents\WEB\AppWebAvKevinCalles\webapppolimusicv1_asp
.\deploy.ps1
```

El script realiza automáticamente:

1. Comprueba que IIS esté activo (e inicia el servicio si hace falta).
2. Habilita la característica **ASP clásico** si no está instalada.
3. Copia la aplicación a `C:\inetpub\wwwroot\webapppolimusicv1_asp`.
4. Concede permisos NTFS a `IIS_IUSRS` (lectura general y **escritura** en     
   `assets\` y `uploads\` para las cargas de archivos).
5. Configura en IIS: `enableParentPaths`, límite de petición de 30 MB y tipos MIME.
6. Verifica el despliegue por HTTP.

Para desinstalar:

```powershell
.\deploy.ps1 -Uninstall
```

### 4.3 Dirección de acceso

```
http://localhost/webapppolimusicv1_asp/
```



---

## 5. Cuentas de prueba

| Usuario      | Contraseña | Rol            |
| ------------ | ---------- | -------------- |
| `userAdmin`  | `12345678` | Administrador  |
| `epn.fis`    | `12345678` | Usuario normal |
| `userNormal` | `12345678` | Usuario normal |

---

## 6. Uso

### Usuario normal

1. Inicie sesión con `epn.fis` / `12345678`.
2. **Canciones** → pulse una fila para reproducirla o use los botones     
   *Reproducir todo* / *Aleatorio*.
3. **Álbumes** y **Artistas** → consulte las fichas y reproduzca sus canciones.
4. **Playlists** → cree, edite, elimine y gestione las canciones de sus listas.
5. Botón **+** de cada canción → agréguela a una de sus playlists.

### Administrador (`userAdmin`)

Además de lo anterior, la barra lateral muestra la sección **Administración**:

| Pantalla  | Operaciones                                                      |
| --------- | ---------------------------------------------------------------- |
| Panel     | Indicadores y estado de los recursos                             |
| Canciones | Crear / editar / eliminar, subir MP3, asociar artistas y álbumes |
| Artistas  | Crear / editar / eliminar, subir fotografía                      |
| Álbumes   | Crear / editar / eliminar, definir la lista de canciones         |
| Playlists | Gestionar las playlists de todos los usuarios                    |
| Usuarios  | Crear / editar / eliminar, asignar rol y restablecer contraseña  |

### Diagnóstico

`http://localhost/webapppolimusicv1_asp/diagnostics.asp`

Comprueba la conexión a SQL Server, la existencia de las 9 tablas, las carpetas de  
recursos, la presencia de los archivos MP3, el formato de las contraseñas y la  
prueba interna de SHA-1. Sin sesión muestra solo la parte técnica, lo que permite  
detectar problemas antes de poder iniciar sesión.

---

## 7. Seguridad implementada

- **Escapado de SQL**: toda entrada se normaliza con `SqlStr()` (duplicación de    
  comillas) o con conversión numérica estricta antes de concatenarse.
- **Salida escapada**: `H()` codifica HTML en todo valor mostrado.
- **Contraseñas**: SHA-1; la comparación se hace en el servidor.
- **Control de acceso**: `RequireLogin()` y `RequireAdmin()` en cada página;    
  `PlaylistCanEdit()` valida la propiedad antes de cualquier modificación.
- **Cargas de archivos**: nombre saneado, extensión validada, límite de tamaño y    
  escritura fuera de la carpeta de código.
- **Redirecciones**: `SafeLocalUrl()` evita redirecciones a sitios externos.
- **Multipart**: se usa `Request.BinaryRead` una sola vez y nunca se accede a    
  `Request.Form` después (evita el error 0106 de ASP).

---

## 8. Notas sobre los recursos

- El paquete `SongFilesAndImages_resources.zip` incluía **un solo MP3**    
  (`bensound-adventure.mp3`) pero la base de datos referencia cinco    
  (`adventure`, `dreams`, `energy`, `scifi`, `tomorrow`). Para que las cinco    
  canciones de ejemplo puedan reproducirse, las cuatro restantes se generaron    
  como **copias del mismo archivo**. Reemplace esos archivos por las pistas    
  reales en `assets/SongFiles/` o súbalos desde    
  *Administración → Canciones*.
- Las imágenes referenciadas en la base (`~/albumImages/...`,    
  `~/ArtistImages/...`, `~/userImages/username1.png`) no venían en el paquete.    
  Cuando un archivo no existe, la aplicación genera un marcador SVG con    
  degradado turquesa e inicial, por lo que la interfaz se ve completa.

---


## 9. Solución de problemas

| Síntoma                                        | Causa probable / solución                                                                                                                         |
| ---------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| `Error de base de datos`                       | El servicio SQL Server está detenido, o la base `BDD_PoliMusic` no existe. Ejecute `BDD_PoliMusic.sql`.                                           |
| `No fue posible conectar`                      | Falta el login `usr_polimusic`. Vuelva a ejecutar el script de la base de datos.                                                                  |
| La página muestra el código ASP en texto plano | La característica **ASP clásico** no está habilitada. Ejecute `deploy.ps1` como administrador.                                                    |
| `Error ASP 0131 Disallowed Parent Path`        | No debería ocurrir (las páginas están en la raíz). Si aparece, habilite rutas padre con `appcmd set config /section:asp /enableParentPaths:true`. |
| No se pueden subir MP3 o imágenes              | El límite de ASP es 200 KB por omisión. Ejecute `deploy.ps1` como administrador o suba los archivos manualmente a `assets/SongFiles/`.            |
| Las canciones suenan iguales                   | Es el comportamiento esperado: solo se recibió un MP3 (ver sección 8).                                                                            |
| Error de acceso denegado al escribir           | Conceda permisos de escritura a `IIS_IUSRS` en `assets\` y `uploads\`.                                                                            |
| Los cambios no se ven                          | Fuerce la recarga con `Ctrl + F5` (el CSS y el JS llevan la versión como parámetro).                                                              |
| Redirige a `/login.asp` en vez de a `/webapppolimusicv1_asp/login.asp` | Era un fallo de `AppBase()` al deducir la raíz virtual; **ya corregido**. Vuelva a desplegar con `desplegar.bat`.                     |
| HTTP 500 en todas las páginas                  | Errores de compilación en `includes/sha1.asp` (sufijo `#`) y `includes/ui.asp` (identificador `sTop`); **ya corregidos**. Vuelva a desplegar.     |

### Incidencias corregidas tras la primera validación en IIS

1. **`includes/sha1.asp` no compilaba.** Los literales usaban el sufijo de tipo
   de VBA (`4294967296#`, `2147483648#`…). VBScript acepta el sufijo en una
   asignación simple, pero **no dentro de una condición `If`**, donde provocaba
   *"Se esperaba 'Then'"*. Se sustituyeron por literales `Double` explícitos
   (`4294967296.0`). El algoritmo no cambia y `tools/sha1_check.py` sigue
   verificándolo contra `hashlib`.
2. **`includes/ui.asp` no compilaba.** La variable `sTop` es inválida en
   VBScript porque en minúsculas coincide con la palabra reservada `stop`.
   Renombrada a `sLimit`. (Se comprobaron los 196 identificadores declarados
   del proyecto contra el motor real: no queda ninguno inválido.)
3. **`AppBase()` deducía mal la raíz virtual.** Usaba `APPL_MD_PATH`, que
   devuelve `/` cuando la carpeta es una simple subcarpeta del sitio y no una
   aplicación IIS. Ahora la raíz se obtiene del directorio de la página en
   ejecución (`SCRIPT_NAME`), lo que funciona tanto en la raíz del sitio como
   en `/webapppolimusicv1_asp/`.

---

## 10. Herramientas de validación (carpeta `tools/`, no se despliega)

| Archivo              | Uso                                                                                                                        |
| -------------------- | -------------------------------------------------------------------------------------------------------------------------- |
| `vbslint.py`         | Valida estructura de bloques y variables no declaradas en el VBScript. `python tools/vbslint.py --all`                     |
| `vbscompile.py`      | Traduce cada `.asp` a un `.vbs` equivalente para compilarlo con el **motor real de VBScript**.                              |
| `vbscheck.ps1`       | Ejecuta la compilación anterior y reporta archivo, línea y mensaje de error.                                                |
| `vbscallcheck.py`    | Verifica que toda función o sub invocada exista en el cierre de inclusiones de cada página. `python tools/vbscallcheck.py` |
| `htmlcheck.py`       | Comprueba el equilibrio de etiquetas HTML en el marcado de cada página.                                                    |
| `sha1_check.py`      | Contrasta la implementación de SHA-1 con `hashlib`. `python tools/sha1_check.py`                                           |
| `sql_validation.sql` | Ejecuta las 20 consultas de la aplicación contra `BDD_PoliMusic`.                                                          |
| `seed_usuarios.sql`  | Crea las cuentas `userAdmin`, `epn.fis` y `userNormal`.                                                                    |
| `reorganize.py`      | Utilidad usada para trasladar las páginas de `admin/` y `api/` a la raíz.                                                  |

### Compilación real del VBScript (sin desplegar)

`tools/vbscompile.py` traduce cada página a VBScript puro y `tools/vbscheck.ps1`
lo compila con `MSScriptControl` — el mismo motor (`vbscript.dll`) que usa IIS —
de modo que los errores de compilación se detectan **antes** de desplegar y con
el número de línea exacto:

```powershell
python tools/vbscompile.py
& "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" -File tools\vbscheck.ps1
```

Se generan dos pruebas por página: el **cierre completo** (página + todos sus
`#include`, que es exactamente lo que compila IIS) y el **archivo suelto** (para
localizar errores dentro de un include concreto).

---

## 11. Modelo de datos utilizado

| Tabla               | Uso en la aplicación                       |
| ------------------- | ------------------------------------------ |
| `TBL_USER`          | Cuentas y roles (`USER_TYPE`)              |
| `TBL_SONG`          | Catálogo de canciones y contador `PLAYS`   |
| `TBL_ARTIST`        | Artistas                                   |
| `TBL_ALBUM`         | Álbumes                                    |
| `TBL_TRACKLIST`     | Relación **álbum ↔ canción**               |
| `TBL_PLAYER`        | Relación **canción ↔ artista**             |
| `TBL_PLAYLIST`      | Playlists (`TYPE` = 0 pública / 1 privada) |
| `TBL_PLAYLIST_SONG` | Relación **playlist ↔ canción**            |
| `TBL_CATALOG`       | Catálogo de tipos (roles y visibilidad)    |

---

**PoliMusic v1.0.0** — ASP Clásico + SQL Server + IIS
