#Requires -RunAsAdministrator
<#
=====================================================================
 PoliMusic - Script de despliegue en IIS
=====================================================================
 Copia la aplicacion ASP clasico a C:\inetpub\wwwroot\webapppolimusicv1_asp,
 habilita ASP clasico si hace falta, ajusta los permisos NTFS y configura
 los limites de ASP.

 Uso (PowerShell como Administrador):

     cd <carpeta del proyecto>
     .\deploy.ps1

 Parametros opcionales:

     .\deploy.ps1 -Target "C:\inetpub\wwwroot\webapppolimusicv1_asp"
     .\deploy.ps1 -Uninstall
=====================================================================
#>
[CmdletBinding()]
param(
    [string]$Source   = '',
    [string]$Target   = '',
    [string]$SiteName = '',
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------
# Resolucion de rutas
# ---------------------------------------------------------------------
# IMPORTANTE: no se usa $PSScriptRoot como valor por omision del parametro.
# En algunas formas de invocacion (por ejemplo desde un .bat elevado con
# -File) $PSScriptRoot llega VACIO en el bloque param(), y entonces
# Test-Path falla con:
#   "No se puede enlazar el argumento con el parametro 'Path' porque es
#    una cadena vacia."
# Se resuelve aqui, con varios respaldos y un mensaje claro si falla.
# ---------------------------------------------------------------------
$scriptDir = ''
if ($PSScriptRoot) {
    $scriptDir = $PSScriptRoot
} elseif ($MyInvocation.MyCommand.Path) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
} elseif ($MyInvocation.MyCommand.Definition) {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition
}
if ([string]::IsNullOrWhiteSpace($scriptDir)) { $scriptDir = (Get-Location).Path }

if ([string]::IsNullOrWhiteSpace($Source))   { $Source   = $scriptDir }
if ([string]::IsNullOrWhiteSpace($Target))   { $Target   = 'C:\inetpub\wwwroot\webapppolimusicv1_asp' }
if ([string]::IsNullOrWhiteSpace($SiteName)) { $SiteName = 'Default Web Site' }

$resolved = Resolve-Path -LiteralPath $Source -ErrorAction SilentlyContinue
if ($resolved) { $Source = $resolved.Path } else { $Source = '' }
if ([string]::IsNullOrWhiteSpace($Source)) {
    throw "No se pudo determinar la carpeta de origen. Ejecute el script desde la carpeta del proyecto o indique la ruta: .\deploy.ps1 -Source 'C:\ruta\a\webapppolimusicv1_asp'"
}

$AppCmd = Join-Path $env:windir 'system32\inetsrv\appcmd.exe'
$tmpDir = $scriptDir
if ($env:TEMP) { $tmpDir = $env:TEMP }
$LogPath = Join-Path $tmpDir 'polimusic_deploy.log'
$Marker  = Join-Path $tmpDir 'polimusic_deploy.done'

try { Start-Transcript -Path $LogPath -Force | Out-Null } catch {
    Write-Host "    [!]  No se pudo iniciar el registro en $LogPath" -ForegroundColor Yellow
}
Remove-Item $Marker -Force -ErrorAction SilentlyContinue

function Write-Step([string]$Text) {
    Write-Host ''
    Write-Host "==> $Text" -ForegroundColor Cyan
}
function Write-Ok([string]$Text)   { Write-Host "    [OK] $Text"   -ForegroundColor Green }
function Write-Warn2([string]$Text){ Write-Host "    [!]  $Text"   -ForegroundColor Yellow }

# ---------------------------------------------------------------------
# Manejo de errores: mensaje claro en lugar del volcado de excepcion
# ---------------------------------------------------------------------
trap {
    Write-Host ''
    Write-Host '=====================================================================' -ForegroundColor Red
    Write-Host ' EL DESPLIEGUE NO PUDO COMPLETARSE' -ForegroundColor Red
    Write-Host '=====================================================================' -ForegroundColor Red
    Write-Host ('  Motivo: ' + $_.Exception.Message) -ForegroundColor Red
    Write-Host ''
    Write-Host '  Compruebe:' -ForegroundColor Yellow
    Write-Host '   1. Que el script se ejecute como Administrador.'
    Write-Host '   2. Que la carpeta de origen sea la del proyecto (contiene default.asp).'
    Write-Host '      Si no se detecta, indiquela explicitamente, por ejemplo:'
    Write-Host '        .\deploy.ps1 -Source "C:\ruta\a\webapppolimusicv1_asp"'
    Write-Host '   3. Que el servicio IIS (W3SVC) este iniciado.'
    Write-Host ''
    try { Stop-Transcript | Out-Null } catch { }
    exit 1
}

# ---------------------------------------------------------------------
# Desinstalacion
# ---------------------------------------------------------------------
if ($Uninstall) {
    Write-Step "Eliminando $Target"
    if (Test-Path $Target) {
        Remove-Item -Path $Target -Recurse -Force
        Write-Ok "Carpeta eliminada."
    } else {
        Write-Warn2 "La carpeta no existia."
    }
    Write-Host ''
    Write-Host 'PoliMusic fue retirado de IIS.' -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------------
# 1. Comprobaciones previas
# ---------------------------------------------------------------------
Write-Step 'Comprobando el entorno'

Write-Host "    Carpeta del script : $scriptDir"
Write-Host "    Origen             : $Source"
Write-Host "    Destino            : $Target"

$svc = Get-Service W3SVC -ErrorAction SilentlyContinue
if (-not $svc) {
    throw 'IIS no esta instalado. Instale el rol "Servidor web (IIS)" antes de continuar.'
}
if ($svc.Status -ne 'Running') {
    Write-Warn2 "El servicio W3SVC esta detenido; se iniciara."
    Start-Service W3SVC
}
Write-Ok "IIS activo (servicio W3SVC: $($svc.Status))."

if (-not (Test-Path -LiteralPath $Source)) {
    throw "No se encuentra la carpeta de origen: '$Source'"
}
if (-not (Test-Path -LiteralPath (Join-Path $Source 'default.asp'))) {
    throw "El origen no parece ser la aplicacion PoliMusic (falta default.asp): '$Source'"
}
Write-Ok "Origen valido: $Source"

# ---------------------------------------------------------------------
# 2. Habilitar ASP clasico
# ---------------------------------------------------------------------
Write-Step 'Verificando la caracteristica "ASP clasico"'

$aspDll = Join-Path $env:windir 'system32\inetsrv\asp.dll'
if (Test-Path $aspDll) {
    Write-Ok 'ASP clasico ya esta instalado.'
} else {
    Write-Warn2 'ASP clasico no esta instalado. Habilitando...'
    $r = Start-Process -FilePath 'dism.exe' `
        -ArgumentList '/online', '/enable-feature', '/featurename:IIS-ASP', '/all', '/norestart' `
        -Wait -PassThru -NoNewWindow
    if ($r.ExitCode -ne 0 -and $r.ExitCode -ne 3010) {
        throw "No fue posible habilitar ASP clasico (dism salio con $($r.ExitCode))."
    }
    Write-Ok 'ASP clasico habilitado.'
}

# ---------------------------------------------------------------------
# 3. Copiar la aplicacion
#     Se copia de forma explicita (solo lo que forma parte de la app) para
#     no arrastrar carpetas de trabajo como tools/ o .workbuddy-ai/.
# ---------------------------------------------------------------------
Write-Step "Copiando la aplicacion a $Target"

if (-not (Test-Path $Target)) {
    New-Item -ItemType Directory -Path $Target -Force | Out-Null
}

# Limpieza de restos de despliegues anteriores
$stray = @('tools', '.workbuddy-ai', '.git', '.vs', 'node_modules', '__pycache__')
foreach ($d in $stray) {
    $p = Join-Path $Target $d
    if (Test-Path $p) {
        Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
        Write-Warn2 "Resto eliminado del destino: $d"
    }
}
Get-ChildItem -Path $Target -File -ErrorAction SilentlyContinue |
    Where-Object { $_.Extension -in @('.sql', '.zip', '.pyc', '.py') } |
    ForEach-Object {
        Remove-Item $_.FullName -Force -ErrorAction SilentlyContinue
        Write-Warn2 "Resto eliminado del destino: $($_.Name)"
    }

# Paginas ASP de la raiz
Copy-Item -Path (Join-Path $Source '*.asp') -Destination $Target -Force
$nAsp = (Get-ChildItem -Path (Join-Path $Source '*.asp') -File).Count
Write-Ok "$nAsp paginas ASP copiadas a la raiz."

# Configuracion y documentacion
foreach ($f in @('web.config', 'README.md')) {
    $p = Join-Path $Source $f
    if (Test-Path $p) { Copy-Item $p -Destination $Target -Force }
}
Write-Ok 'web.config y README.md copiados.'

# Libreria compartida y recursos. Se copia el CONTENIDO en modo combinacion
# (sin borrar el destino) para no perder los archivos que el usuario haya
# subido desde el panel: assets/SongFiles, assets/UserImages, etc.
foreach ($d in @('includes', 'assets')) {
    $src = Join-Path $Source $d
    if (Test-Path $src) {
        $dst = Join-Path $Target $d
        if (-not (Test-Path $dst)) {
            New-Item -ItemType Directory -Path $dst -Force | Out-Null
        }
        Copy-Item -Path (Join-Path $src '*') -Destination $dst -Recurse -Force
        $n = (Get-ChildItem -Path $dst -Recurse -File).Count
        Write-Ok "$d\ copiada ($n archivos en el destino)."
    }
}

# Carpeta de cargas
$uploads = Join-Path $Target 'uploads'
if (-not (Test-Path $uploads)) { New-Item -ItemType Directory -Path $uploads -Force | Out-Null }
Write-Ok 'uploads\ preparada.'

# ---------------------------------------------------------------------
# 4. Permisos NTFS
# ---------------------------------------------------------------------
Write-Step 'Ajustando permisos NTFS'

$principal = 'IIS_IUSRS'
& icacls $Target /grant "${principal}:(OI)(CI)RX" /T /Q | Out-Null
Write-Ok "Lectura y ejecucion concedidas a $principal."

# La aplicacion escribe archivos subidos en assets/ y uploads/
foreach ($writable in @('assets', 'uploads')) {
    $p = Join-Path $Target $writable
    if (Test-Path $p) {
        & icacls $p /grant "${principal}:(OI)(CI)M" /T /Q | Out-Null
        Write-Ok "Escritura concedida en $writable\ (cargas de archivos)."
    }
}

# ---------------------------------------------------------------------
# 5. Configuracion de ASP en IIS
# ---------------------------------------------------------------------
Write-Step 'Configurando ASP en IIS'

if (Test-Path $AppCmd) {
    $site = & $AppCmd list site "$SiteName" 2>$null
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($site)) {
        Write-Warn2 "No se encontro el sitio '$SiteName'; se omite la comprobacion por sitio."
    } else {
        Write-Ok "Sitio encontrado: $SiteName"
    }

    # Detalle de errores de script en el navegador (util en desarrollo;
    # en produccion conviene ponerlo en false).
    & $AppCmd set config /section:asp /scriptErrorSentToBrowser:true 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'scriptErrorSentToBrowser = true (errores detallados).' }

    # Rutas padre: no son necesarias (todas las paginas estan en la raiz y
    # usan includes/...), pero se activan por robustez.
    & $AppCmd set config /section:asp /enableParentPaths:true 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'enableParentPaths = true.' }

    # Limite del cuerpo de la peticion: 30 MB (para subir MP3 e imagenes)
    & $AppCmd set config /section:asp /limits.maxRequestEntityAllowed:31457280 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'maxRequestEntityAllowed = 30 MB.' }

    & $AppCmd set config /section:requestFiltering /requestLimits.maxAllowedContentLength:31457280 2>$null | Out-Null
    if ($LASTEXITCODE -eq 0) { Write-Ok 'maxAllowedContentLength = 30 MB.' }

    # Tipos MIME que el filtro podria bloquear
    foreach ($ext in @('.mp3', '.svg', '.webp')) {
        & $AppCmd set config /section:staticContent /+"[fileExtension='$ext',mimeType='application/octet-stream']" 2>$null | Out-Null
    }
    Write-Ok 'Tipos MIME para .mp3 / .svg / .webp verificados.'
} else {
    Write-Warn2 'No se encontro appcmd.exe; configure ASP manualmente en el Administrador de IIS.'
}

# ---------------------------------------------------------------------
# 6. Verificacion por HTTP
# ---------------------------------------------------------------------
Write-Step 'Verificando el despliegue'

$base = 'http://localhost/webapppolimusicv1_asp/'
$checks = @(
    @{ Url = $base;                Expect = '302' },
    @{ Url = "${base}login.asp";   Expect = '200' },
    @{ Url = "${base}default.asp"; Expect = '302' },
    @{ Url = "${base}placeholder.asp"; Expect = '200' }
)
$allOk = $true
foreach ($c in $checks) {
    try {
        $resp = Invoke-WebRequest -Uri $c.Url -UseBasicParsing -TimeoutSec 25 `
            -MaximumRedirection 0 -ErrorAction Stop
        $code = [string]$resp.StatusCode
    } catch {
        $code = ''
        if ($_.Exception.Response) { $code = [string][int]$_.Exception.Response.StatusCode }
    }
    if ($code -eq $c.Expect -or ($c.Expect -eq '302' -and $code -eq '200')) {
        Write-Ok "$($c.Url) -> HTTP $code"
    } else {
        $allOk = $false
        Write-Warn2 "$($c.Url) -> HTTP $code (se esperaba $($c.Expect))"
    }
}

Write-Host ''
if ($allOk) {
    Write-Host '=====================================================================' -ForegroundColor Green
    Write-Host ' PoliMusic desplegado correctamente' -ForegroundColor Green
    Write-Host '=====================================================================' -ForegroundColor Green
} else {
    Write-Host '=====================================================================' -ForegroundColor Yellow
    Write-Host ' Despliegue completado con advertencias' -ForegroundColor Yellow
    Write-Host '=====================================================================' -ForegroundColor Yellow
    Write-Host ' Revise el Administrador de IIS y la seccion de diagnostico.'
}
Write-Host ''
Write-Host "  Aplicacion  : $base"
Write-Host "  Diagnostico : ${base}diagnostics.asp"
Write-Host "  Carpeta     : $Target"
Write-Host ''
Write-Host '  Cuentas de prueba:'
Write-Host '    userAdmin  / 12345678   (administrador)'
Write-Host '    epn.fis    / 12345678   (usuario normal)'
Write-Host '    userNormal / 12345678   (usuario normal)'
Write-Host ''
Write-Host "  Registro de este despliegue: $LogPath"
Write-Host ''

'OK' | Set-Content -Path $Marker -Encoding ASCII
try { Stop-Transcript | Out-Null } catch { }
