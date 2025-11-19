@echo off
echo ========================================
echo Registro Alternativo del Add-in Azurriga
echo ========================================
echo.
echo Este metodo registra el add-in usando un catalogo compartido.
echo Es mas confiable que el metodo de carpeta Wef.
echo.
pause

echo.
echo INSTRUCCIONES:
echo.
echo 1. Abre Excel
echo.
echo 2. Ve a: Archivo ^> Opciones ^> Centro de confianza ^> 
echo    Configuracion del Centro de confianza
echo.
echo 3. Click en "Catalogos de complementos de confianza"
echo.
echo 4. En "Direccion URL del catalogo", pega esta URL:
echo.
echo    https://albertoalgora.github.io/excel-addin-azurriga/
echo.
echo 5. Marca la casilla "Mostrar en menu"
echo.
echo 6. Click en "Agregar catalogo" y luego "Aceptar"
echo.
echo 7. Reinicia Excel
echo.
echo 8. Ve a: Insertar ^> Complementos de Office ^> CARPETA COMPARTIDA
echo.
echo 9. Busca "Add-in Azurriga" y click en Agregar
echo.
echo.

echo Copiando la URL al portapapeles...
echo https://albertoalgora.github.io/excel-addin-azurriga/ | clip
echo.
echo La URL ha sido copiada al portapapeles!
echo Puedes pegarla directamente con Ctrl+V
echo.
pause
