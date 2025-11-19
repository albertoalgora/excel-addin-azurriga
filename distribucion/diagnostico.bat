@echo off
echo ========================================
echo Diagnostico de Instalacion del Add-in
echo ========================================
echo.

echo Verificando instalacion de Office...
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport
echo.

echo Verificando carpeta Wef...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef" (
    echo [OK] Carpeta Wef existe
    echo.
    echo Manifests instalados:
    dir /b "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\*.xml"
    echo.
) else (
    echo [ERROR] Carpeta Wef no existe
    echo.
)

echo Verificando manifest de Azurriga...
if exist "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\manifest-azurriga.xml" (
    echo [OK] Manifest encontrado
    echo.
    echo Contenido del SourceLocation:
    type "%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\manifest-azurriga.xml" | findstr "SourceLocation"
    echo.
) else (
    echo [ERROR] Manifest no encontrado
    echo.
)

echo Verificando acceso a GitHub Pages...
echo Probando conexion...
powershell -Command "try { $response = Invoke-WebRequest -Uri 'https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml' -UseBasicParsing; Write-Host '[OK] GitHub Pages accesible'; Write-Host 'Status:' $response.StatusCode } catch { Write-Host '[ERROR] No se puede acceder a GitHub Pages'; Write-Host $_.Exception.Message }"
echo.

echo ========================================
echo Presiona una tecla para continuar...
pause >nul
