@echo off
chcp 65001 > nul
cd /d "%~dp0"
echo Ejecutando script de recopilacion OTD Americas...
powershell.exe -ExecutionPolicy Bypass -NoProfile -File "%~dp0Recopilar-DatosEquipo.ps1"
echo.
echo Listo. Puedes cerrar esta ventana y cargar el JSON en la pagina.
pause