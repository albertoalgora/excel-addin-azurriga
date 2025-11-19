@echo off
echo ========================================
echo Instalador de Add-in Azurriga para Excel
echo (Version de Produccion - GitHub Pages)
echo ========================================
echo.
echo Este script registrara el add-in desde GitHub Pages.
echo NO requiere servidor local ni npm.
echo.
echo Requisitos:
echo - Microsoft Excel 2016 o superior
echo - Conexion a Internet
echo.
pause

REM Verificar si Excel esta instalado
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: No se detecto Microsoft Office instalado.
    echo Por favor, instala Microsoft Excel e intenta de nuevo.
    echo.
    pause
    exit /b 1
)

REM Crear carpeta de manifests si no existe
if not exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" mkdir "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef"

REM Limpiar manifests antiguos para evitar conflictos
echo Limpiando instalaciones previas...
del /q "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.xml" 2>nul

REM Crear archivo manifest temporal que apunta a GitHub Pages
echo Instalando el add-in...

REM Descargar manifest de produccion desde GitHub
powershell -Command "& {Invoke-WebRequest -Uri 'https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml' -OutFile '%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\manifest-azurriga.xml'}"

if %errorlevel% neq 0 (
    echo.
    echo ERROR: No se pudo descargar el manifest desde GitHub Pages.
    echo Verifica tu conexion a Internet.
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Instalacion completada exitosamente!
echo ========================================
echo.
echo El add-in se ha registrado correctamente.
echo.
echo PROXIMOS PASOS:
echo.
echo 1. CIERRA Excel si esta abierto (completamente)
echo.
echo 2. Abre Excel de nuevo
echo.
echo 3. Ve a: Insertar -^> Complementos de Office -^> MIS COMPLEMENTOS
echo    (o en la pestana Desarrollador si la tienes habilitada)
echo.
echo 4. Busca "Add-in Azurriga para Excel" en la lista
echo.
echo 5. Haz click en el add-in para abrirlo
echo.
echo NOTA: El add-in funciona desde la nube (GitHub Pages)
echo       No necesitas ningun servidor local activo.
echo.
pause
