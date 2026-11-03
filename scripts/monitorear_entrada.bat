@echo off
setlocal
cd /d "%~dp0.."
if not exist ".venv\Scripts\python.exe" (
  echo Ejecutar scripts\instalar.bat primero.
  pause
  exit /b 1
)
"%CD%\.venv\Scripts\python.exe" automatizar_fichadas.py --monitorear "%CD%\data\entrada"
pause
