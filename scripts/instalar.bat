@echo off
setlocal
cd /d "%~dp0.."
where py >nul 2>&1
if errorlevel 1 (
  echo [ERROR] No se encontro el lanzador de Python 'py'.
  echo Instalar Python 3.10 o superior con el lanzador para Windows.
  pause
  exit /b 1
)
py -m venv .venv
if errorlevel 1 goto :error
"%CD%\.venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error
if not exist "data\entrada" mkdir "data\entrada"
if not exist "data\salidas_fichadas" mkdir "data\salidas_fichadas"
if not exist "data\historico_originales" mkdir "data\historico_originales"
echo Instalacion completada.
pause
exit /b 0
:error
echo Error de instalacion. Revisar conexion a internet y Python.
pause
exit /b 1
