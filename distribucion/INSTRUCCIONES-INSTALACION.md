# 📦 Instalación de Add-in Azurriga para Excel

**Versión de Producción** - Compatible con cualquier equipo con Excel

---

## ✅ Requisitos
- Microsoft Excel (versión 2016 o superior, Desktop o Web)
- Conexión a Internet
- Credenciales de acceso al servidor Azurriga

---

## 🚀 Instalación Rápida (Recomendado)

### Pasos:

1. **Abre Excel** (Desktop o Web en https://office.com)

2. **Ve al menú de Complementos:**
   - Excel Desktop: **Inicio → Complementos** (o **Insertar → Complementos de Office**)
   - Excel Web: **Inicio → Complementos**

3. **Busca "Complementos de Office"** y haz click

4. **Selecciona "MIS COMPLEMENTOS"** (arriba)

5. **Carga el add-in desde URL:**
   - Click en **"CARGAR MI COMPLEMENTO"** (esquina superior)
   - Pega esta URL:
     ```
     https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml
     ```
   - Click en **"Cargar"**

6. **¡Listo!** El add-in aparecerá en el panel lateral derecho de Excel

---

## 📝 Método Alternativo: Instalación con Catálogo (Excel Desktop)

### Pasos:

1. **Abre Excel** y ve a **Archivo → Opciones**

2. **Centro de confianza → Configuración del Centro de confianza**

3. Click en **Catálogos de complementos de confianza**

4. En **"Dirección URL del catálogo"**, pega:
   ```
   https://albertoalgora.github.io/excel-addin-azurriga/
   ```

5. Marca la casilla **"Mostrar en menú"**

6. Click en **"Agregar catálogo"** → **Aceptar** → **Aceptar**

7. **Reinicia Excel**

8. Ve a **Inicio → Complementos → Complementos de Office → CARPETA COMPARTIDA**

9. Busca **"Add-in Azurriga"** y click en **Agregar**

---

## 🔐 Primer Uso: Login

Al abrir el add-in por primera vez:

1. **Se abrirá un panel lateral** a la derecha con el logo de Azurriga

2. **Click en el botón "Login"**

3. **Introduce tus credenciales:**
   - Usuario: Tu usuario del sistema Azurriga
   - Contraseña: Tu contraseña del sistema

4. **Click en "Iniciar Sesión"**

5. Si las credenciales son correctas, verás una notificación verde de **"Login exitoso"**

6. Los botones **"Download"** e **"Import"** se habilitarán

---

## 📊 Uso: Descargar Datos

Una vez autenticado:

1. **Click en "Download"**

2. **Selecciona el tipo de datos:**
   - **Cuentas** - Listado de cuentas contables
   - **Flujos** - Códigos de flujo de caja
   - **Códigos Presupuestarios** - Códigos presupuestarios (Code, Id, Description)
   - **Movimientos** - Transacciones de caja (requiere configuración adicional)

3. **Elige cantidad de registros:**
   - 50, 100, 500 o Todos

4. **Para Movimientos (opcional):**
   - Selecciona una cuenta específica
   - Filtra por rango de fechas
   - Marca los campos que deseas ver

5. **Click en "Descargar"**

6. **Los datos aparecerán automáticamente** en una nueva hoja de Excel con:
   - Encabezados formateados
   - Fechas en formato correcto
   - Columnas autoajustadas

---

## ❓ Solución de Problemas

### No veo el menú "Complementos"

**Habilita la pestaña Desarrollador:**
1. Excel → Archivo → Opciones → Personalizar cinta
2. Marca la casilla **"Desarrollador"**
3. Aceptar
4. Ahora verás: **Desarrollador → Complementos de Office**

### Error: "No se puede cargar el complemento"

1. **Verifica tu conexión a Internet**
2. **Cierra completamente Excel:**
   - Task Manager (Ctrl+Shift+Esc) → Busca "Excel" → Finalizar tarea
3. **Abre Excel de nuevo** e intenta cargar otra vez

### Error de CORS o "blocked by CORS policy"

- ✅ **Este error ya está solucionado** en la versión actual
- Si lo ves, asegúrate de que Excel esté conectado a Internet
- Cierra y abre Excel de nuevo

### Error: "Cannot find module '@libsql/client'"

- ⚠️ **Problema conocido en Vercel**: El logging está temporalmente deshabilitado
- El add-in funciona correctamente - este error no afecta su funcionamiento
- Estamos trabajando en una solución alternativa para el logging en producción

### Error al hacer Login: "Credenciales incorrectas"

1. Verifica que tu usuario y contraseña sean correctos
2. El sistema es sensible a mayúsculas/minúsculas
3. No incluyas espacios extras al principio o final

### El add-in se cierra solo o desaparece

- Excel puede desactivar add-ins que tardan mucho en cargar
- Ve a **Archivo → Opciones → Complementos → Administrar: Complementos deshabilitados**
- Si ves "Add-in Azurriga", selecciónalo y click en **"Habilitar"**

---

## 🔄 Desinstalación

Si deseas desinstalar el add-in:

### Excel Desktop:
1. Ve a **Inicio → Complementos → Complementos de Office**
2. Click derecho en **"Add-in Azurriga"** → **Quitar**

### Desinstalación completa:
1. Cierra Excel
2. Abre el Explorador de archivos
3. Ve a: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`
4. Elimina todos los archivos `.xml` relacionados con Azurriga
5. Reinicia Excel

---

## 📞 Información Técnica

### URLs del Proyecto:
- **Manifest XML:** https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml
- **Aplicación Web:** https://albertoalgora.github.io/excel-addin-azurriga/
- **Proxy API:** https://excel-addin-azurriga.vercel.app/api/proxy
- **Servidor OData:** https://azprod.azurriga.com:1035/

### Arquitectura:
```
Excel (GitHub Pages) ←→ Proxy (Vercel) ←→ Servidor OData (Azurriga)
```

### Soporte:
- **Política de privacidad:** https://albertoalgora.github.io/excel-addin-azurriga/privacy-policy.html
- **Términos de uso:** https://albertoalgora.github.io/excel-addin-azurriga/terms-of-use.html

---

## ⚙️ Configuración Avanzada (Opcional)

### Usar en múltiples equipos:

El add-in funciona en **cualquier equipo** sin instalación adicional:
- Los datos se guardan en la **nube de Office 365**
- Solo necesitas cargar el manifest XML una vez por cuenta de usuario
- Tu sesión (credenciales) **NO** se guarda - debes iniciar sesión cada vez

### Compartir con tu equipo:

1. Comparte la URL del manifest:
   ```
   https://albertoalgora.github.io/excel-addin-azurriga/manifest-production.xml
   ```

2. Cada usuario debe:
   - Cargar el add-in siguiendo las instrucciones de instalación
   - Iniciar sesión con sus propias credenciales
   - Los datos descargados quedan en su propio Excel

### Seguridad:

- ✅ Conexión HTTPS segura
- ✅ Autenticación Basic Auth sobre HTTPS
- ✅ Las credenciales **NO** se almacenan en el navegador
- ✅ Cada sesión es independiente
- ✅ Los datos solo se descargan cuando el usuario lo solicita

---

**Última actualización:** Enero 2026  
**Versión:** 1.0 (Producción estable)
