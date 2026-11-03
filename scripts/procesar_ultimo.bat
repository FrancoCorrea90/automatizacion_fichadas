@echo off
setlocal EnableExtensions

REM Ubicarse en la raiz del proyecto
cd /d "%~dp0.."

REM Python del entorno virtual
set "PYTHON=%CD%\.venv\Scripts\python.exe"

if not exist "%PYTHON%" (
    echo [ERROR] No se encontro el entorno virtual.
    echo Ejecutar scripts\instalar.bat antes de continuar.
    pause
    exit /b 1
)

REM Carpeta donde RRHH coloca los archivos
set "ENTRADA=%CD%\data\entrada"

if not exist "%ENTRADA%" mkdir "%ENTRADA%"

REM Buscar el Excel mas reciente
set "ULTIMO="

for /f "usebackq delims=" %%F in (`powershell -NoProfile -Command "$f = Get-ChildItem -LiteralPath '%ENTRADA%' -File | Where-Object { $_.Extension -in '.xls','.xlsx' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1; if ($f) { $f.Name }"`) do (
    set "ULTIMO=%%F"
)

REM Comprobar si existe un archivo
if not defined ULTIMO (
    echo [ERROR] No hay archivos Excel en data\entrada.
    pause
    exit /b 1
)

echo ============================================
echo    Automatizacion de fichadas
echo ============================================
echo.
echo Procesando: %ULTIMO%
echo.

REM Ejecutar el script principal
"%PYTHON%" automatizar_fichadas.py --archivo "%ENTRADA%\%ULTIMO%"

set "CODIGO=%ERRORLEVEL%"

echo.
echo Codigo de salida: %CODIGO%
pause

exit /b %CODIGO%