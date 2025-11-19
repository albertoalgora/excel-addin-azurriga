@echo off
echo ========================================
echo Cargar Add-in directamente en Excel 2019
echo ========================================
echo.
echo Este metodo carga el manifest directamente usando VBA.
echo.
pause

REM Descargar manifest a una ubicación temporal
set MANIFEST_PATH=%TEMP%\manifest-azurriga.xml
echo Descargando manifest...
powershell -Command "& {Invoke-WebRequest -Uri 'https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml' -OutFile '%MANIFEST_PATH%'}"

if %errorlevel% neq 0 (
    echo ERROR: No se pudo descargar el manifest.
    pause
    exit /b 1
)

echo.
echo ========================================
echo INSTRUCCIONES PARA EXCEL 2019:
echo ========================================
echo.
echo 1. Abre Excel
echo.
echo 2. Presiona Alt + F11 (abre el Editor de VBA)
echo.
echo 3. En el menu: Ver ^> Ventana Inmediato (o Ctrl + G)
echo.
echo 4. Copia y pega este comando en la ventana inmediato:
echo.
echo    Application.COMAddIns.Add "%MANIFEST_PATH%"
echo.
echo 5. Presiona Enter
echo.
echo 6. Cierra el Editor VBA (Alt + Q)
echo.
echo 7. El add-in deberia aparecer en la pestaña Inicio
echo.
echo.
echo ALTERNATIVA - Habilitar pestaña Desarrollador:
echo.
echo 1. Archivo ^> Opciones ^> Personalizar cinta de opciones
echo.
echo 2. Marca la casilla "Desarrollador" en el panel derecho
echo.
echo 3. Click Aceptar
echo.
echo 4. Ve a: Desarrollador ^> Complementos de COM ^> Agregar
echo.
echo 5. Busca el archivo: %MANIFEST_PATH%
echo.
echo.
pause
