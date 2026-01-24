# 📘 Información del Proyecto - Add-in Azurriga para Excel

## 📊 Descripción General

**Add-in Azurriga para Excel** es un complemento de Microsoft Excel que permite conectarse a un servidor OData para descargar datos contables (cuentas, flujos de caja, movimientos contables) de forma rápida y automática con autenticación segura y formato automático de datos.

**Proveedor**: Azurriga  
**Versión**: 1.0.0.0  
**Locales**: es-ES  

---

## 🏗️ Arquitectura del Proyecto

### Estructura General

```
Excel Add-in (Office.js)
         ↓
    HTTPS Connection
         ↓
Webpack Dev Server (Dev) / GitHub Pages (Prod) / Vercel Proxy
         ↓
    HTTP Connection
         ↓
OData Server (azprod.azurriga.com:1035)
```

### Tecnologías Principales

- **Office.js**: API oficial de Microsoft para Office Add-ins
- **Webpack 5**: Empaquetador de módulos y servidor de desarrollo
- **Babel**: Transpilador ES6+
- **Core-js**: Polyfills para compatibilidad
- **HTML5 + CSS3**: Interfaz de usuario
- **Fetch API**: Comunicación HTTP
- **OData v4**: Protocolo de comunicación con backend
- **Turso DB (@libsql/client)**: Base de datos para logging
- **Express.js**: Servidor proxy local
- **Vercel**: Serverless Functions para proxy HTTPS

---

## 📁 Estructura del Proyecto

```
excel-addin-azurriga/
├── src/                              # Código fuente
│   ├── taskpane/                     # Panel de tareas del add-in
│   │   ├── taskpane.html             # UI principal
│   │   ├── taskpane.js               # Lógica de negocio (868 líneas)
│   │   └── taskpane.css              # Estilos
│   └── commands/                     # Comandos de Excel ribbon
│       ├── commands.html
│       └── commands.js
│
├── api/                              # Funciones serverless (Vercel)
│   ├── proxy.js                      # Proxy HTTPS → HTTP OData
│   ├── db.js                         # Cliente Turso DB para logging
│   └── stats.js                      # Estadísticas de uso
│
├── azure-proxy/                      # Azure Functions (alternativa)
│   ├── ODataProxy/
│   │   ├── index.js
│   │   └── function.json
│   ├── host.json
│   └── package.json
│
├── assets/                           # Recursos gráficos
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
│
├── distribucion/                     # Paquete de instalación
│   ├── instalar-addin-produccion.bat
│   ├── manifest-production.xml
│   ├── INSTRUCCIONES-INSTALACION.md
│   └── iconos/
│
├── docs/                             # Documentación
│   ├── DOCUMENTACION-TECNICA.md      # Documentación técnica completa
│   ├── DATABASE.md
│   └── Documentacion-AddIn-Azurriga.html
│
├── manifest.xml                      # Manifiesto de desarrollo
├── manifest-production.xml           # Manifiesto de producción
├── webpack.config.js                 # Configuración de Webpack
├── babel.config.json                 # Configuración de Babel
├── package.json                      # Dependencias y scripts
├── vercel.json                       # Configuración de Vercel
├── staticwebapp.config.json          # Configuración de Azure Static Web Apps
├── dev-proxy-server.js               # Servidor proxy local
└── VERCEL-DEPLOYMENT.md              # Guía de despliegue en Vercel
```

---

## 🌐 URLs y Endpoints

### Entornos

#### **Desarrollo (Local)**
- **Dev Server**: `https://localhost:3000/`
- **Taskpane**: `https://localhost:3000/taskpane.html`
- **Commands**: `https://localhost:3000/commands.html`
- **Proxy Local**: `/odata/` (proxy de webpack)

#### **Producción**
- **GitHub Pages**: `https://albertoalgora.github.io/excel-addin-azurriga/`
- **Vercel Proxy**: `https://excel-addin-azurriga.vercel.app/`
- **Soporte**: `https://albertoalgora.github.io/excel-addin-azurriga/support.html`

### Servidor OData (Backend)

- **Servidor Principal**: `https://azprod.azurriga.com:1035/`
- **Endpoint Base**: `/odata/`
- **Autenticación**: HTTP Basic Authentication

### Endpoints del Proxy Vercel

```
Base URL: https://excel-addin-azurriga.vercel.app/api/

/api/proxy?path=odata/AccountSet           # Descargar cuentas
/api/proxy?path=odata/CashflowSet          # Descargar flujos de caja
/api/proxy?path=odata/JournalEntrySet      # Descargar movimientos
/api/stats                                  # Estadísticas de uso
```

#### Ejemplo de uso:
```javascript
fetch('https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/AccountSet?$top=50', {
  headers: {
    'Authorization': 'Basic ' + btoa(username + ':' + password)
  }
})
```

---

## ⚙️ Configuraciones Importantes

### 1. Webpack Configuration (`webpack.config.js`)

```javascript
const urlDev = "https://localhost:3000/";
const urlProd = "https://www.contoso.com/"; // Cambiar en producción

devServer: {
  port: 3000,
  https: true,
  proxy: {
    '/odata': {
      target: 'https://azprod.azurriga.com:1035',
      secure: false,
      changeOrigin: true
    }
  }
}
```

### 2. Vercel Configuration (`vercel.json`)

```json
{
  "functions": {
    "api/proxy.js": {
      "maxDuration": 10
    }
  }
}
```

### 3. CORS Configuration (`staticwebapp.config.json`)

```json
{
  "globalHeaders": {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET, POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type, Authorization"
  }
}
```

### 4. Manifest Configuration

**ID del Add-in**: `adc7c00e-0ac0-4f01-ab64-4defa945ac7a`

**Desarrollo** (`manifest.xml`):
- SourceLocation: `https://localhost:3000/taskpane.html`
- IconUrl: `https://localhost:3000/assets/icon-32.png`

**Producción** (`manifest-production.xml`):
- SourceLocation: `https://albertoalgora.github.io/excel-addin-azurriga/taskpane.html`
- IconUrl: `https://albertoalgora.github.io/excel-addin-azurriga/assets/icon-32.png`

---

## 🔧 Scripts NPM Disponibles

### Desarrollo
```bash
npm run dev-server          # Inicia el servidor de desarrollo (puerto 3000)
npm run build:dev           # Build de desarrollo con source maps
npm run watch               # Build con watch mode
npm start                   # Inicia Office Add-in debugging en Excel Desktop
```

### Producción
```bash
npm run build               # Build de producción optimizado
npm run lint                # Ejecuta ESLint
npm run lint:fix            # Ejecuta ESLint y arregla problemas
npm run validate            # Valida el manifest.xml
```

### Debugging
```bash
npm start -- desktop --app excel        # Debug en Excel Desktop
npm start -- desktop --app word         # Debug en Word Desktop
npm start -- desktop --app outlook      # Debug en Outlook Desktop
npm start -- desktop --app powerpoint   # Debug en PowerPoint Desktop
npm stop                                # Detiene la sesión de debugging
```

### Testing
```bash
node test-db.js                # Test de conexión a Turso DB
node test-proxy.js             # Test del proxy
node test-count-records.js     # Test de conteo de registros
node test-schema.js            # Test del esquema de base de datos
node check-logs.js             # Ver logs básicos
node check-registros.js        # Ver registros resumidos
node check-registros-detallado.js  # Ver registros detallados
```

---

## 🚀 Flujo de Trabajo de Desarrollo

### 1. Configurar el entorno local

```bash
# Clonar el repositorio
git clone https://github.com/albertoalgora/excel-addin-azurriga.git
cd excel-addin-azurriga

# Instalar dependencias
npm install

# Instalar certificados de desarrollo
npx office-addin-dev-certs install
```

### 2. Iniciar el servidor de desarrollo

```bash
npm run dev-server
```

El servidor estará disponible en `https://localhost:3000/`

### 3. Cargar el add-in en Excel

**Opción A - Usar comando NPM**:
```bash
npm start
```

**Opción B - Sideload manual**:
1. Abrir Excel
2. Ir a **Insertar → Mis complementos → Cargar complemento**
3. Seleccionar `manifest.xml`

### 4. Debugging

- **Chrome DevTools**: F12 en el panel de tareas del add-in
- **VS Code**: Configuración de launch.json disponible
- **Logs**: Console.log se muestra en DevTools

### 5. Build de producción

```bash
npm run build
```

Los archivos se generan en la carpeta `dist/`

---

## 📦 Despliegue

### Opción 1: GitHub Pages (Recomendado)

1. Hacer push a la rama `main`
2. Configurar GitHub Pages desde Settings → Pages
3. Rama: `gh-pages` o `main` (con carpeta `dist/`)
4. La URL será: `https://albertoalgora.github.io/excel-addin-azurriga/`

### Opción 2: Vercel (Para Proxy)

```bash
# Instalar Vercel CLI
npm i -g vercel

# Desplegar
vercel

# Producción
vercel --prod
```

**URL resultante**: `https://excel-addin-azurriga.vercel.app/`

Ver `VERCEL-DEPLOYMENT.md` para más detalles.

### Opción 3: Azure Functions

La carpeta `azure-proxy/` contiene una implementación alternativa usando Azure Functions.

```bash
cd azure-proxy
npm install
func start    # Desarrollo local
```

### Opción 4: IIS (Windows Server)

```powershell
# Ejecutar como administrador
.\setup-iis-site.ps1
```

---

## 🔐 Autenticación y Seguridad

### Flujo de Autenticación

1. Usuario ingresa credenciales en el modal de login
2. Se crea un header `Authorization: Basic <base64(username:password)>`
3. Se hace una petición de prueba a `/odata/AccountSet?$top=1`
4. Si la respuesta es 200 OK, las credenciales se guardan en memoria
5. Las credenciales se envían en cada petición subsecuente

### Características de Seguridad

- ✅ **HTTPS end-to-end** en producción
- ✅ **HTTP Basic Authentication** con el servidor OData
- ✅ **No se almacenan credenciales** en disco (solo en memoria)
- ✅ **CORS configurado** para dominios específicos
- ✅ **Proxy serverless** evita problemas de Mixed Content
- ✅ **Certificados SSL** automáticos en Vercel

### Credenciales en Desarrollo vs Producción

**Desarrollo**:
```javascript
const isDevelopment = window.location.hostname === 'localhost';
const baseUrl = isDevelopment 
  ? '/odata/AccountSet'  // Proxy de webpack
  : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/AccountSet';
```

---

## 📊 Funcionalidades del Add-in

### 1. Login
- Modal de autenticación
- Validación de credenciales contra OData
- Estado de sesión persistente en memoria

### 2. Descarga de Datos

**Tipos de datos disponibles**:
- **Cuentas** (`AccountSet`)
- **Flujos de Caja** (`CashflowSet`)
- **Movimientos Contables** (`JournalEntrySet`)

**Opciones de descarga**:
- Límite de registros: 50, 100, 200, 500, 1000, Todos
- Filtros OData: `$top`, `$filter`, `$orderby`
- Para movimientos: opción de filtrar por fechas

### 3. Formato Automático
- Detección automática de columnas numéricas, de fecha y de texto
- Aplicación de formato Excel apropiado
- Encabezados en negrita
- Ancho de columnas automático

### 4. Importación de Datos
- Descarga en segundo plano
- Inserción en la hoja activa
- Barra de progreso visual
- Manejo de errores robusto

---

## 🗄️ Base de Datos (Turso DB)

### Configuración

El proyecto usa **Turso DB** (base de datos SQLite distribuida) para logging.

**Variables de entorno requeridas**:
```
TURSO_DATABASE_URL=libsql://...
TURSO_AUTH_TOKEN=...
```

### Schema

```sql
CREATE TABLE request_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
  method TEXT,
  path TEXT,
  status_code INTEGER,
  response_time INTEGER,
  user_agent TEXT,
  referer TEXT,
  ip_address TEXT,
  error_message TEXT
);
```

### Scripts de Consulta

```bash
node check-logs.js                    # Ver últimos 50 logs
node check-registros.js               # Resumen de registros
node check-registros-detallado.js     # Detalle completo
```

---

## 🛠️ Troubleshooting

### Problemas Comunes

#### 1. "Add-in no se carga"
- Verificar que el servidor de desarrollo esté corriendo: `npm run dev-server`
- Verificar certificados HTTPS: `npx office-addin-dev-certs install`
- Limpiar caché de Office: Cerrar Excel completamente y volver a abrir

#### 2. "Mixed Content error"
- En producción, usar siempre el proxy de Vercel
- Verificar que todas las URLs usen HTTPS

#### 3. "CORS error"
- Verificar configuración de `staticwebapp.config.json`
- Verificar headers CORS en el proxy

#### 4. "Authentication failed"
- Verificar credenciales
- Verificar que el servidor OData esté accesible
- Ver logs en DevTools Console

#### 5. "Build failed"
- Limpiar node_modules: `rm -rf node_modules && npm install`
- Verificar versión de Node.js: >= 14.x recomendado
- Verificar webpack.config.js

### Logs y Debugging

**En el navegador**:
- F12 → Console para ver logs de JavaScript
- F12 → Network para ver peticiones HTTP
- F12 → Sources para debugging con breakpoints

**En Vercel**:
- Dashboard → Proyecto → Functions → View Logs
- Logs en tiempo real de las peticiones al proxy

**Localmente**:
```bash
# Ver logs de webpack
npm run dev-server

# Ver logs de Node.js (proxy local)
node dev-proxy-server.js
```

---

## 📚 Documentación Adicional

- **Documentación Técnica Completa**: [docs/DOCUMENTACION-TECNICA.md](docs/DOCUMENTACION-TECNICA.md)
- **Guía de Instalación**: [distribucion/INSTRUCCIONES-INSTALACION.md](distribucion/INSTRUCCIONES-INSTALACION.md)
- **Despliegue en Vercel**: [VERCEL-DEPLOYMENT.md](VERCEL-DEPLOYMENT.md)
- **Base de Datos**: [docs/DATABASE.md](docs/DATABASE.md)
- **Soporte**: https://albertoalgora.github.io/excel-addin-azurriga/support.html

---

## 🔗 Enlaces Importantes

- **Repositorio GitHub**: https://github.com/albertoalgora/excel-addin-azurriga
- **GitHub Pages (Prod)**: https://albertoalgora.github.io/excel-addin-azurriga/
- **Vercel Proxy**: https://excel-addin-azurriga.vercel.app/
- **Office.js Docs**: https://learn.microsoft.com/en-us/office/dev/add-ins/
- **OData v4**: https://www.odata.org/documentation/

---

## 👥 Equipo y Contacto

**Desarrollador**: Alberto Algora  
**Organización**: Azurriga  
**Soporte**: https://albertoalgora.github.io/excel-addin-azurriga/support.html

---

## 📝 Licencia

MIT License - Ver repositorio para detalles completos.

---

**Última actualización**: Enero 2026  
**Versión del documento**: 1.0
