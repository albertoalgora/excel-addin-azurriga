@echo off
echo ========================================
echo Configuracion de Catalogo Local
echo ========================================
echo.
echo Este script crea un catalogo local para el add-in.
echo.
pause

REM Crear carpeta para el catálogo
set CATALOG_PATH=%USERPROFILE%\Documents\OfficeAddInsCatalog
if not exist "%CATALOG_PATH%" mkdir "%CATALOG_PATH%"

echo Descargando manifest al catalogo local...
powershell -Command "& {Invoke-WebRequest -Uri 'https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml' -OutFile '%CATALOG_PATH%\manifest-azurriga.xml'}"

if %errorlevel% neq 0 (
    echo.
    echo ERROR: No se pudo descargar el manifest.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Catalogo creado en:
echo %CATALOG_PATH%
echo ========================================
echo.
echo AHORA SIGUE ESTOS PASOS:
echo.
echo 1. Abre Excel
echo.
echo 2. Ve a: Archivo ^> Opciones ^> Centro de confianza ^> 
echo    Configuracion del Centro de confianza
echo.
echo 3. Click en "Catalogos de complementos de confianza"
echo.
echo 4. En "Direccion URL del catalogo", pega:
echo.

REM Convertir ruta a formato de red
echo    file:///%CATALOG_PATH:\=/%
echo.
echo 5. Marca "Mostrar en menu"
echo.
echo 6. Click en "Agregar catalogo" y "Aceptar"
echo.
echo 7. Reinicia Excel
echo.
echo 8. Ve a: Insertar ^> Complementos ^> MIS COMPLEMENTOS
echo.
echo 9. Deberia aparecer "Add-in Azurriga" en CARPETA COMPARTIDA
echo.

echo Copiando la ruta al portapapeles...
echo file:///%CATALOG_PATH:\=/% | clip
echo.
echo [OK] Ruta copiada! Usa Ctrl+V para pegarla.
echo.
pause
