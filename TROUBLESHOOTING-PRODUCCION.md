# 🔧 Guía de Diagnóstico - Add-in en Producción

## Problema Reportado
El complemento falla al inicio de la carga en Excel de escritorio usando la versión de producción.

## Pasos de Diagnóstico

### 1. Verificar qué error específico se muestra

**En Excel Desktop:**
1. Abre el panel de tareas del add-in
2. Presiona **F12** para abrir las Herramientas de Desarrollador
3. Ve a la pestaña **Console**
4. Busca mensajes de error en rojo

**Errores comunes:**
- `Failed to load resource: net::ERR_CONNECTION_REFUSED` → El servidor no está disponible
- `Office is not defined` → Office.js no se cargó correctamente
- `CORS error` → Problema de política de origen cruzado
- `404 Not Found` → Archivo no encontrado en GitHub Pages

---

### 2. Verificar la URL del manifiesto

El manifiesto de producción apunta a:
```
https://albertoalgora.github.io/excel-addin-azurriga/taskpane.html
```

**Verificar en el navegador:**
1. Abre esta URL en Chrome/Edge: https://albertoalgora.github.io/excel-addin-azurriga/taskpane.html
2. ¿Se carga la página correctamente?
3. Abre F12 → Console → ¿Hay errores?

---

### 3. Verificar que GitHub Pages esté actualizado

Los archivos en `dist/` necesitan estar sincronizados con GitHub Pages.

**Pasos:**
1. Ve al repositorio en GitHub: https://github.com/albertoalgora/excel-addin-azurriga
2. Entra a la carpeta donde están los archivos (probablemente raíz o carpeta `dist/`)
3. Verifica la fecha de última actualización de `taskpane.html` y `taskpane.js`
4. ¿Son del 16 de febrero de 2026 o posteriores?

---

### 4. Verificar configuración de GitHub Pages

1. Ve a **Settings** → **Pages** en el repositorio
2. Verifica que esté configurado como:
   - **Source**: Deploy from a branch
   - **Branch**: `main` (o la rama que uses)
   - **Folder**: `/` (root) o `/dist`

---

### 5. Desplegar archivos actualizados a GitHub Pages

Si los archivos en GitHub no están actualizados, necesitas subirlos:

```powershell
# Desde la carpeta del proyecto
cd c:\Desarrollo\Azurriga\excel-addin-azurriga

# Compilar archivos de producción (ya lo hiciste)
npm run build

# Commit y push al repositorio
git add dist/
git commit -m "Actualizar archivos de producción - Fix error de inicialización"
git push origin main
```

Si GitHub Pages está configurado en carpeta raíz (no `/dist`), necesitas copiar los archivos:

```powershell
# Copiar archivos de dist/ a la raíz
Copy-Item -Path "dist\*" -Destination "." -Recurse -Force

# Commit y push
git add .
git commit -m "Actualizar archivos de producción en raíz"
git push origin main
```

---

### 6. Limpiar caché de Office

El problema podría ser que Excel está usando archivos antiguos en caché:

1. **Cerrar completamente Excel** (Archivo → Salir o Alt+F4)
2. **Limpiar caché de Office:**
   ```powershell
   # Ejecutar en PowerShell como administrador
   Remove-Item -Path "$env:LOCALAPPDATA\Microsoft\Office\16.0\Wef\*" -Recurse -Force
   ```
3. **Reiniciar Excel** y volver a cargar el add-in

---

### 7. Reinstalar el complemento

1. En Excel: **Insertar → Mis complementos**
2. Click derecho en "Add-in Azurriga" → **Quitar**
3. **Cerrar Excel completamente**
4. Usar el script de instalación:
   ```powershell
   cd c:\Desarrollo\Azurriga\excel-addin-azurriga\distribucion
   .\instalar-addin-produccion.bat
   ```
5. Abrir Excel y verificar

---

### 8. Verificar conectividad con el proxy de Vercel

El add-in necesita conectarse a:
```
https://excel-addin-azurriga.vercel.app/api/proxy
```

**Probar en navegador:**
```
https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/$metadata
```

Deberías recibir un error 401 (Unauthorized) o 200, pero NO un error de conexión.

---

### 9. Probar con versión local primero

Para verificar que el código funciona, prueba en desarrollo:

```powershell
# Terminal 1 - Iniciar servidor de desarrollo
npm run dev-server

# Terminal 2 - Cargar add-in en Excel
npm start
```

Si funciona en local pero no en producción, el problema es de despliegue/GitHub Pages.

---

## Soluciones Rápidas

### Solución A: Forzar recarga sin caché
En Excel, con el panel del add-in abierto:
- Presiona **Ctrl + Shift + R** (o **Ctrl + F5**)
- Cierra y vuelve a abrir Excel

### Solución B: Usar manifest absoluto
Asegurarse de que `manifest-production.xml` tiene URLs completas:

```xml
<SourceLocation DefaultValue="https://albertoalgora.github.io/excel-addin-azurriga/taskpane.html"/>
<IconUrl DefaultValue="https://albertoalgora.github.io/excel-addin-azurriga/assets/icon-32.png"/>
```

### Solución C: Verificar Office.js está cargándose
Agregar un script de verificación en taskpane.html (temporal para debug):

```html
<script>
console.log("HTML cargado - verificando Office.js");
setTimeout(() => {
  if (typeof Office === 'undefined') {
    console.error("ERROR: Office.js no está disponible después de 2 segundos");
  } else {
    console.log("OK: Office.js está disponible");
  }
}, 2000);
</script>
```

---

## Información de Contacto

Si el problema persiste después de seguir estos pasos, recopilar:

1. **Screenshot del error** en la consola de Chrome DevTools (F12)
2. **Mensaje de error exacto** que muestra Excel
3. **Versión de Excel**: Archivo → Cuenta → Acerca de Excel
4. **URL que está intentando cargar** el add-in

---

## Checklist de Verificación

- [ ] Los archivos en `dist/` están actualizados (16 de febrero o después)
- [ ] Los archivos se subieron a GitHub (`git push`)
- [ ] GitHub Pages está activo y desplegado
- [ ] La URL `https://albertoalgora.github.io/excel-addin-azurriga/taskpane.html` carga en navegador
- [ ] F12 en la URL anterior no muestra errores
- [ ] Caché de Office limpiado
- [ ] Excel completamente cerrado y reabierto
- [ ] Add-in reinstalado con `instalar-addin-produccion.bat`
- [ ] Vercel proxy responde: `https://excel-addin-azurriga.vercel.app/`
