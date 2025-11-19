# 📦 Instalación de Add-in Azurriga para Excel

## ✅ Requisitos
- Microsoft Excel (versión 2016 o superior)
- Conexión a Internet

## 📥 Método 1: Instalación Automática (Windows)

### Pasos:

1. **Descarga este paquete completo** (carpeta con todos los archivos)

2. **Ejecuta el archivo `instalar-addin-produccion.bat`**
   - Click derecho → Ejecutar como administrador (si es necesario)
   - Sigue las instrucciones en pantalla

3. **Cierra Excel** si está abierto

4. **Abre Excel de nuevo**

5. **Busca el add-in:**
   - Opción A: **Insertar → Complementos de Office → MIS COMPLEMENTOS**
   - Opción B: **Desarrollador → Complementos de Office** (si tienes esa pestaña)
   - Opción C: En la cinta, busca un icono de **"Complementos"**

6. Deberías ver **"Add-in Azurriga para Excel"** en la lista

---

## 📝 Método 2: Instalación Manual

### Para Excel Desktop:

1. Descarga el archivo `manifest-production.xml`

2. Abre **Excel**

3. Ve a **Archivo → Opciones → Centro de confianza → Configuración del Centro de confianza**

4. Click en **Catálogos de complementos de confianza**

5. En **"Dirección URL del catálogo"**, pega:
   ```
   https://albertoalgora.github.io/excel-addin-azurriga/
   ```

6. Marca la casilla **"Mostrar en menú"**

7. Click en **Agregar catálogo** → **Aceptar**

8. Reinicia Excel

9. Ve a **Insertar → Complementos de Office → CARPETA COMPARTIDA**

10. Busca **"Add-in Azurriga"** y click en **Agregar**

---

### Para Excel Online (Web):

1. Abre https://www.office.com e inicia sesión

2. Abre **Excel Online** (crea un libro nuevo o abre uno existente)

3. Click en **Insertar** (puede estar en el menú "..." si la ventana es pequeña)

4. Click en **Complementos de Office**

5. Click en **MÁS COMPLEMENTOS**

6. Selecciona la pestaña **CARGAR MI COMPLEMENTO**

7. Pega esta URL:
   ```
   https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml
   ```

8. Click en **Cargar**

---

## ❓ Solución de Problemas

### No veo el menú "Complementos de Office"

**Habilita la pestaña Desarrollador:**
1. Excel → Archivo → Opciones → Personalizar cinta
2. Marca la casilla **"Desarrollador"** en el panel derecho
3. Aceptar
4. Ahora verás: **Desarrollador → Complementos de Office**

### El add-in no se carga

1. Verifica tu conexión a Internet
2. Cierra completamente Excel (Task Manager → Cerrar todos los procesos de Excel)
3. Abre Excel de nuevo
4. Intenta cargar el add-in otra vez

### Error de seguridad o certificado

- El add-in está alojado en GitHub Pages (HTTPS seguro)
- Si ves advertencias, es normal en primera carga
- Click en "Confiar" o "Permitir"

### Error: "Este complemento ya no está disponible"

Este error aparece cuando se mezclan instalaciones de desarrollo y producción:

1. **Cierra Excel completamente**
2. **Elimina todos los manifests antiguos:**
   - Abre el Explorador de archivos
   - Pega en la barra de direcciones: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
   - Elimina todos los archivos `.xml` que encuentres
3. **Ejecuta de nuevo** `instalar-addin-produccion.bat`
4. **Reinicia Excel**

---

## 📞 Soporte

- **Documentación completa:** https://albertoalgora.github.io/excel-addin-azurriga/support.html
- **Política de privacidad:** https://albertoalgora.github.io/excel-addin-azurriga/privacy-policy.html
- **Términos de uso:** https://albertoalgora.github.io/excel-addin-azurriga/terms-of-use.html

---

## 🚀 Uso del Add-in

Una vez instalado:

1. Abre el add-in desde **Inicio → Complementos → Add-in Azurriga**

2. **Inicia sesión** con tus credenciales del servidor OData

3. **Descarga datos:**
   - Click en "Download"
   - Selecciona tipo: Cuentas, Flujos o Movimientos
   - Elige el límite de registros
   - Click en "Descargar"

4. Los datos aparecerán automáticamente en una hoja nueva de Excel

---

## 🔄 Desinstalación

Si deseas desinstalar el add-in:

1. Cierra Excel
2. Ve a: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
3. Elimina el archivo `manifest-azurriga.xml`
4. Reinicia Excel
