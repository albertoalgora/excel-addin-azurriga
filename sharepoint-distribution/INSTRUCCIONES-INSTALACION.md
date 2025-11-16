# Instrucciones de Instalación - Add-in Azurriga para Excel
## Distribución mediante SharePoint App Catalog

---

## 📦 Para el Administrador de SharePoint

### Paso 1: Verificar/Crear el App Catalog

1. Ir al **Centro de Administración de SharePoint**:
   - URL: `https://[tuempresa]-admin.sharepoint.com`

2. En el menú lateral: **Más características** → **Aplicaciones** → **Abrir**

3. Click en **Catálogo de aplicaciones**

4. Si no existe, crear uno nuevo:
   - Click en **Crear un nuevo sitio de catálogo de aplicaciones**
   - Seguir el asistente de configuración
   - **Nota**: Solo se crea una vez por toda la organización

---

### Paso 2: Subir el Add-in al Catálogo

1. Ir al sitio del App Catalog:
   - URL típica: `https://[tuempresa].sharepoint.com/sites/appcatalog`

2. Click en la biblioteca **Aplicaciones para Office**

3. Click en **Nuevo** → **Elegir archivos**

4. Subir el archivo: **manifest-production.xml**

5. Completar el formulario que aparece:
   - **Nombre**: Add-in Azurriga para Excel
   - **Descripción**: Conecta Excel con servidor OData para descargar cuentas, flujos de caja y movimientos contables
   - **Icono**: Subir `icon-128.png` (incluido en esta carpeta)
   - **Categoría**: Productividad / Finanzas
   - **Habilitado**: ✅ Sí
   - **Visibilidad**: Todos los sitios (o específicos según política de la organización)

6. Guardar

---

### Paso 3: Aprobar el Add-in

1. En la biblioteca **Aplicaciones para Office**, localizar el add-in recién subido

2. Click en los tres puntos (...) → **Propiedades**

3. Cambiar el estado a: **Aprobado**

4. Guardar cambios

✅ **El add-in ya está disponible para todos los usuarios de la organización**

---

### Paso 4: (Opcional) Configurar Permisos

Si deseas limitar quién puede usar el add-in:

1. En las propiedades del add-in, configurar:
   - **Sitios específicos**: Listar URLs de sitios SharePoint autorizados
   - **Grupos de usuarios**: Especificar grupos de seguridad

---

## 👥 Para los Usuarios Finales

### Instalar el Add-in desde Excel Desktop

1. Abrir **Microsoft Excel** (versión de escritorio)

2. Ir a la pestaña **Insertar**

3. Click en **Complementos** → **Obtener complementos**

4. En la ventana que se abre, click en la pestaña **MI ORGANIZACIÓN**

5. Buscar **"Add-in Azurriga"** o **"Azurriga"**

6. Click en **Agregar** o **Confiar**

7. El add-in aparecerá en la pestaña **Inicio** con el botón **"Show Task Pane"**

---

### Instalar el Add-in desde Excel Online (Office 365)

1. Abrir **Excel Online** (office.com/launch/excel)

2. Abrir cualquier libro o crear uno nuevo

3. Click en **Insertar** → **Complementos**

4. Click en **MI ORGANIZACIÓN**

5. Buscar **"Add-in Azurriga"**

6. Click en **Agregar**

7. El add-in aparecerá en la cinta de opciones

---

## 🔧 Uso del Add-in

### Primera vez:

1. Click en el botón del add-in en la pestaña **Inicio**

2. Se abrirá un panel lateral

3. Click en **Iniciar sesión**

4. Introducir credenciales del servidor OData:
   - **Usuario**: Tu nombre de usuario
   - **Contraseña**: Tu contraseña

5. Una vez autenticado, seleccionar el tipo de datos a descargar:
   - 📊 **Cuentas**: Código, descripción y saldo
   - 💰 **Flujos de caja**: Descripción, importe y tipo
   - 📝 **Movimientos**: Fecha, descripción, importe, cuenta

6. Los datos se insertarán automáticamente en la hoja activa con formato de tabla

---

## 📋 Requisitos Técnicos

### Para la Organización:
- Licencia de Microsoft 365 con SharePoint Online
- Permisos de administrador de SharePoint
- App Catalog configurado

### Para los Usuarios:
- Microsoft Excel 2016 o posterior, O Excel Online (Office 365)
- Conexión a Internet
- Credenciales válidas del servidor OData

---

## 🌐 URLs y Recursos

- **Servidor de datos**: http://8cf33ac.online-server.cloud:1031/odata/
- **Documentación completa**: https://albertoalgora.github.io/excel-addin-azurriga/Documentacion-AddIn-Azurriga.html
- **Soporte técnico**: https://albertoalgora.github.io/excel-addin-azurriga/support.html
- **Política de privacidad**: https://albertoalgora.github.io/excel-addin-azurriga/privacy-policy.html
- **Términos de uso**: https://albertoalgora.github.io/excel-addin-azurriga/terms-of-use.html
- **Código fuente**: https://github.com/albertoalgora/excel-addin-azurriga

---

## 📧 Contacto y Soporte

**Email de soporte**: soporte@azurriga.com

**Para reportar errores**: https://github.com/albertoalgora/excel-addin-azurriga/issues

**Proveedor**: Azurriga

---

## 🔐 Seguridad y Privacidad

- ✅ Todas las comunicaciones usan **HTTPS**
- ✅ Las credenciales **NO se almacenan** permanentemente
- ✅ Los datos descargados permanecen **solo en el archivo Excel local**
- ✅ No se comparten datos con terceros
- ✅ Cumple con **RGPD** (Reglamento General de Protección de Datos)

---

## 🆘 Solución de Problemas Comunes

### El add-in no aparece en "MI ORGANIZACIÓN"
- Verificar que el add-in está **Aprobado** en el App Catalog
- Esperar hasta 24 horas para propagación en toda la organización
- Cerrar y volver a abrir Excel

### No puedo iniciar sesión
- Verificar credenciales con el administrador
- Comprobar conexión a Internet
- Verificar que el servidor OData está disponible

### Los datos no se insertan
- Verificar que la hoja no está protegida
- Seleccionar una celda específica antes de descargar
- Comprobar permisos de escritura en el archivo

---

## 📦 Archivos Incluidos en esta Carpeta

```
sharepoint-distribution/
├── manifest-production.xml   (Archivo principal del add-in)
├── icon-128.png              (Icono grande para el catálogo)
├── icon-64.png               (Icono mediano)
├── icon-32.png               (Icono pequeño)
└── INSTRUCCIONES-INSTALACION.md (Este archivo)
```

---

## ✅ Checklist de Instalación

### Para el Administrador:
- [ ] Verificar que existe el App Catalog en SharePoint
- [ ] Subir `manifest-production.xml` a "Aplicaciones para Office"
- [ ] Completar información del add-in (nombre, descripción, icono)
- [ ] Cambiar estado a "Aprobado"
- [ ] (Opcional) Configurar permisos específicos
- [ ] Notificar a los usuarios que el add-in está disponible

### Para los Usuarios:
- [ ] Abrir Excel (Desktop u Online)
- [ ] Ir a Insertar → Complementos → MI ORGANIZACIÓN
- [ ] Buscar "Add-in Azurriga" y hacer click en "Agregar"
- [ ] Iniciar sesión con credenciales del servidor
- [ ] ¡Empezar a usar el add-in!

---

**Versión del Add-in**: 1.0.0.0  
**Fecha**: Noviembre 2025  
**Última actualización de este documento**: 12/11/2025
