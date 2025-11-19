@echo off
echo ========================================
echo Instalador de Add-in Azurriga para Excel
echo ========================================
echo.
echo Este script registrará el add-in en tu sistema.
echo.
pause

REM Crear carpeta de manifests si no existe
if not exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" mkdir "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

REM Copiar manifest
echo Copiando manifest...
copy /Y "%~dp0sharepoint-distribution\manifest-production.xml" "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\manifest-azurriga.xml"

echo.
echo ========================================
echo Instalación completada!
echo ========================================
echo.
echo IMPORTANTE: Cierra Excel si está abierto y ábrelo de nuevo.
echo.
echo Luego ve a: Insertar -^> Complementos de Office -^> MIS COMPLEMENTOS
echo.
pause
