# =====================================================================
#  PoliMusic - Compilacion real del VBScript con MSScriptControl
# ---------------------------------------------------------------------
#  Compila con el mismo motor (vbscript.dll) que usa IIS los archivos
#  .vbs generados por tools/vbscompile.py y reporta los errores de
#  COMPILACION con archivo, linea y mensaje.
#
#  Los errores de EJECUCION que aparecen aqui son normales: el codigo se
#  ejecuta fuera de IIS, sin Response/Request/Session reales.
#
#  Uso (PowerShell de 32 bits, que es donde esta registrado el control):
#
#      python tools\vbscompile.py
#      & "$env:windir\SysWOW64\WindowsPowerShell\v1.0\powershell.exe" `
#            -NoProfile -ExecutionPolicy Bypass -File tools\vbscheck.ps1
# =====================================================================
$ErrorActionPreference = 'Continue'

$dir      = Join-Path $env:TEMP 'pmcompile'
$manifest = Join-Path $dir 'manifest.txt'
$out      = Join-Path $env:TEMP 'vbs_compile_result.txt'
$rows     = @()

if (-not (Test-Path $manifest)) {
    Write-Host "No existe $manifest. Ejecute primero: python tools\vbscompile.py" -ForegroundColor Yellow
    exit 1
}

$rows += "=== Compilacion con el motor real de VBScript (MSScriptControl) ==="
$rows += ""

$total = 0
$compilaMal = 0

foreach ($line in Get-Content $manifest) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $parts = $line -split "`t"
    $label = $parts[0]
    $file  = $parts[1]
    if (-not (Test-Path $file)) { continue }
    $total++

    $code = Get-Content -Path $file -Raw -Encoding UTF8
    $sc = $null
    try {
        $sc = New-Object -ComObject MSScriptControl.ScriptControl
        $sc.Language = 'VBScript'
        $sc.AllowUI = $false
        $sc.AddCode($code)
        $rows += ("OK     " + $label)
    } catch {
        $msg = $_.Exception.Message
        $ln = ''
        try { $ln = $sc.Error.Line } catch { }
        $col = ''
        try { $col = $sc.Error.Column } catch { }
        # Solo son errores de compilacion los que el compilador detecta
        # antes de ejecutar nada.
        $esCompilacion = $false
        foreach ($p in @('Se esperaba', 'Error de sintaxis', 'Carácter no válido',
                         'Nombre redefinido', 'no declarada')) {
            if ($msg.StartsWith($p) -or $msg.Contains($p)) { $esCompilacion = $true }
        }
        if ($esCompilacion) { $compilaMal++ }
        $rows += ("FALLA  " + $label)
        $rows += ("       tipo    : " + $(if ($esCompilacion) { 'COMPILACION' } else { 'ejecucion (compila bien)' }))
        $rows += ("       mensaje : " + $msg)
        $rows += ("       linea   : " + $ln + "  columna: " + $col)
    } finally {
        if ($sc -ne $null) {
            try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sc) | Out-Null } catch { }
        }
    }
}

$rows += ""
$rows += ("TOTAL: " + $total + " archivos compilados, " + $compilaMal + " con ERROR DE COMPILACION")
$rows | Set-Content $out -Encoding UTF8

Write-Host ""
if ($compilaMal -eq 0) {
    Write-Host ("OK - " + $total + " archivos compilados sin errores de compilacion.") -ForegroundColor Green
} else {
    Write-Host ("ATENCION - " + $compilaMal + " archivos con error de COMPILACION.") -ForegroundColor Red
    Get-Content $out | Where-Object { $_ -match 'FALLA|tipo    : COMPILACION|mensaje|linea' }
}
Write-Host ("Detalle completo: " + $out)
exit $(if ($compilaMal -eq 0) { 0 } else { 1 })
