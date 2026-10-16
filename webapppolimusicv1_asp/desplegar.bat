@echo off
REM =====================================================================
REM  PoliMusic - Despliegue en IIS con un solo clic
REM
REM  Haga doble clic en este archivo. Windows solicitara permisos de
REM  administrador (UAC) y a continuacion se ejecutara deploy.ps1.
REM =====================================================================
setlocal
cd /d "%~dp0"

net session >nul 2>&1
if %errorlevel% NEQ 0 (
    echo.
    echo  Solicitando permisos de administrador...
    echo.
    powershell -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo  =====================================================================
echo   PoliMusic - Despliegue en IIS
echo  =====================================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0deploy.ps1" -Source "%~dp0."
set RC=%errorlevel%

echo.
if "%RC%"=="0" (
    echo  Despliegue finalizado correctamente.
) else (
    echo  ATENCION: el despliegue termino con codigo %RC%.
    echo  Revise los mensajes anteriores.
)
echo.
pause
endlocal
