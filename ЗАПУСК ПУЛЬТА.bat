@echo off
cd /d "%~dp0"
echo Ischu Python. Esli ego net, sam postavlju. Ne zakryvaj okno.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0zapusk.ps1"
if errorlevel 1 (
  echo.
  echo Ne zapustilos. Otkroj zapusk.log v etoj zhe papke.
  pause
)
exit /b 0
