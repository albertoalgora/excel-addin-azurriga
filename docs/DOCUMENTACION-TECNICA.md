# Documentación Técnica - Add-in Azurriga para Excel

## 📋 Índice
1. [Arquitectura General](#arquitectura-general)
2. [Flujo de Inicio de la Aplicación](#flujo-de-inicio)
3. [Proceso de Autenticación](#proceso-de-autenticación)
4. [Descarga de Datos](#descarga-de-datos)
5. [Métodos y Funciones Principales](#métodos-principales)
6. [Estructura de Archivos](#estructura-de-archivos)
7. [Configuración de Webpack](#configuración-webpack)

---

## 🏗️ Arquitectura General

### Componentes Principales

```
┌─────────────────────────────────────────────────────────┐
│                    Excel Application                     │
│  ┌───────────────────────────────────────────────────┐ │
│  │           Office.js Runtime Environment           │ │
│  │  ┌─────────────────────────────────────────────┐ │ │
│  │  │        Add-in Taskpane (HTML/CSS/JS)        │ │ │
│  │  │                                              │ │ │
│  │  │  • taskpane.html (UI)                       │ │ │
│  │  │  • taskpane.js (Lógica de negocio)         │ │ │
│  │  │  • taskpane.css (Estilos)                  │ │ │
│  │  └─────────────────────────────────────────────┘ │ │
│  └───────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
                            │
                            │ HTTPS
                            ▼
           ┌──────────────────────────────────┐
           │    Webpack Dev Server (Dev)      │
           │    GitHub Pages (Production)     │
           │    localhost:3000 / GitHub.io    │
           └──────────────────────────────────┘
                            │
                            │ HTTP (Proxy en Dev)
                            ▼
           ┌──────────────────────────────────┐
           │      OData Server (Backend)      │
           │  8cf33ac.online-server.cloud     │
           │         Puerto 1031              │
           └──────────────────────────────────┘
```

### Tecnologías Utilizadas
- **Office.js**: API oficial de Microsoft para Office Add-ins
- **Webpack 5**: Empaquetador de módulos
- **Babel**: Transpilador para JavaScript moderno
- **Core-js**: Polyfills para compatibilidad con navegadores antiguos
- **HTML5 + CSS3**: Interfaz de usuario
- **Fetch API**: Comunicación con el servidor OData
- **OData v4**: Protocolo de comunicación con el backend

---

## 🚀 Flujo de Inicio de la Aplicación

### 1. Carga Inicial del Add-in

**Archivo**: `src/taskpane/taskpane.js`

```javascript
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").classList.add("hidden");
    document.getElementById("app-body").classList.remove("hidden");
    
    // Agregar event listeners para los botones
    document.getElementById("login").onclick = login;
    document.getElementById("download").onclick = showDownloadModal;
    document.getElementById("import").onclick = importData;
    
    // Event listener para cambio de tipo de descarga
    document.getElementById("downloadType").onchange = function() {
      const movimientosOptions = document.getElementById("movimientosOptions");
      if (this.value === "movimientos") {
        movimientosOptions.classList.remove("hidden");
      } else {
        movimientosOptions.classList.add("hidden");
      }
    };
  }
});
```

**¿Qué hace este método?**

1. **Office.onReady()**: Función proporcionada por Office.js
   - Se ejecuta cuando el entorno de Office está listo
   - Recibe un objeto `info` con información sobre la aplicación host

2. **Verificación del host**:
   - `info.host === Office.HostType.Excel` verifica que estamos en Excel
   - Evita errores si el add-in se carga en otra aplicación de Office

3. **Manipulación del DOM**:
   - Oculta el mensaje de carga inicial (`sideload-msg`)
   - Muestra el cuerpo principal de la aplicación (`app-body`)

4. **Registro de Event Listeners**:
   - **Botón Login**: `login()` - Abre modal de autenticación
   - **Botón Download**: `showDownloadModal()` - Abre modal de descarga
   - **Botón Import**: `importData()` - Importa datos (función de ejemplo)
   - **Select downloadType**: Muestra/oculta opciones específicas de movimientos

**Estado inicial de la aplicación**:
```javascript
let userCredentials = {
  username: null,
  password: null,
  isLoggedIn: false
};
```

---

## 🔐 Proceso de Autenticación

### 2. Método `login()`

**Función**: Gestiona todo el proceso de autenticación del usuario

**Flujo detallado**:

```
Usuario hace click en "Iniciar sesión"
    │
    ▼
login() se ejecuta
    │
    ├─► Muestra modal de login
    │   (loginModal)
    │
    ├─► Usuario introduce credenciales
    │   (username, password)
    │
    ├─► Click en "Iniciar sesión" (loginSubmit)
    │   │
    │   ├─► Validación de campos vacíos
    │   │
    │   ├─► Muestra spinner de carga
    │   │   (loginLoading)
    │   │
    │   ├─► Desactiva botones temporalmente
    │   │   (submitButton, cancelButton)
    │   │
    │   ├─► Crea header de autenticación básica
    │   │   Base64(username:password)
    │   │
    │   ├─► Realiza petición al servidor
    │   │   fetch('/odata/')
    │   │   │
    │   │   ├─► ✅ Respuesta exitosa (200 OK)
    │   │   │   │
    │   │   │   ├─► Guarda credenciales
    │   │   │   │   userCredentials = {...}
    │   │   │   │
    │   │   │   ├─► Actualiza UI del botón login
    │   │   │   │   "¡Bienvenido [username]!"
    │   │   │   │   Color verde (#107C10)
    │   │   │   │
    │   │   │   ├─► Activa botones Download e Import
    │   │   │   │   classList.remove("is-disabled")
    │   │   │   │
    │   │   │   ├─► Cierra modal
    │   │   │   │
    │   │   │   └─► Muestra notificación de éxito
    │   │   │
    │   │   └─► ❌ Respuesta fallida (401, 403, 500...)
    │   │       │
    │   │       ├─► Lee mensaje de error del servidor
    │   │       │
    │   │       ├─► Muestra error en el modal
    │   │       │   (loginError)
    │   │       │
    │   │       └─► Auto-oculta error después de 5s
    │   │
    │   └─► Catch: Error de red o conexión
    │       │
    │       ├─► Analiza tipo de error
    │       │   • Failed to fetch: Error de conexión
    │       │   • NetworkError: Sin internet
    │       │
    │       ├─► Muestra mensaje detallado
    │       │
    │       └─► Auto-oculta después de 7s
    │
    └─► Botón "Cancelar" cierra el modal
```

**Código clave**:

```javascript
// Crear header de autenticación básica
const authString = btoa(username + ':' + password);

// Petición con proxy (evita CORS y Mixed Content)
const response = await fetch('/odata/', {
  method: 'GET',
  headers: {
    'Authorization': `Basic ${authString}`,
    'Content-Type': 'application/json',
  }
});
```

**Detalles técnicos**:

1. **Autenticación Básica HTTP**:
   - `btoa()`: Codifica en Base64 la cadena `username:password`
   - Header: `Authorization: Basic [base64]`
   - Método estándar soportado por OData

2. **Uso del Proxy**:
   - URL: `/odata/` (relativa)
   - Webpack dev server redirige a: `http://8cf33ac.online-server.cloud:1031/odata/`
   - Resuelve problemas de:
     - **CORS**: Cross-Origin Resource Sharing
     - **Mixed Content**: HTTPS → HTTP

3. **Gestión de estado**:
   ```javascript
   userCredentials = {
     username: "usuario_ingresado",
     password: "contraseña_ingresada",
     isLoggedIn: true
   };
   ```

---

## 📥 Descarga de Datos

### 3. Método `showDownloadModal()`

**Función**: Abre el modal de configuración de descarga

```javascript
export async function showDownloadModal() {
  try {
    // Verificar que el usuario esté logueado
    if (!userCredentials.isLoggedIn) {
      showNotification("Debe iniciar sesión primero", "error");
      return;
    }

    const modal = document.getElementById("downloadModal");
    modal.classList.remove("hidden");
    modal.style.display = "block";

    // Configurar botón de submit
    document.getElementById("downloadSubmit").onclick = async () => {
      await executeDownload();
    };

    // Configurar botón de cancelar
    document.getElementById("downloadCancel").onclick = () => {
      modal.classList.add("hidden");
    };
  } catch (error) {
    console.error("Error al abrir modal de descarga:", error);
    showNotification("Error al abrir el modal de descarga", "error");
  }
}
```

**¿Qué hace?**

1. **Verificación de autenticación**:
   - Comprueba `userCredentials.isLoggedIn`
   - Si no está autenticado, muestra error y sale

2. **Muestra el modal**:
   - Contiene:
     - Select de tipo de descarga (Cuentas, Flujos, Movimientos)
     - Límite de registros (50, 100, 500, Todos)
     - Campos seleccionables (solo para Movimientos)

3. **Configura event listeners**:
   - Botón "Descargar": Ejecuta `executeDownload()`
   - Botón "Cancelar": Cierra el modal

---

### 4. Método `executeDownload()`

**Función**: Recopila las opciones seleccionadas y ejecuta la descarga

```javascript
async function executeDownload() {
  try {
    const downloadType = document.getElementById("downloadType").value;
    const recordLimit = document.getElementById("recordLimit").value;
    
    // Recoger campos seleccionados para Movimientos
    let selectedFields = [];
    if (downloadType === "movimientos") {
      const checkboxes = document.querySelectorAll('#movimientosOptions input[type="checkbox"]:checked');
      selectedFields = Array.from(checkboxes).map(cb => cb.value);
      
      if (selectedFields.length === 0) {
        showNotification("Debe seleccionar al menos un campo", "error");
        return;
      }
    }

    // Cerrar el modal
    document.getElementById("downloadModal").classList.add("hidden");

    // Llamar a la función de descarga con los parámetros
    await download(downloadType, recordLimit, selectedFields);
  } catch (error) {
    console.error("Error en executeDownload:", error);
    showNotification("Error al preparar la descarga", "error");
  }
}
```

**¿Qué hace?**

1. **Recopila valores del formulario**:
   - **downloadType**: "cuentas" | "flujos" | "movimientos"
   - **recordLimit**: "50" | "100" | "500" | "all"
   - **selectedFields**: Array de campos seleccionados (solo para movimientos)

2. **Validación específica para Movimientos**:
   - Si no se selecciona ningún campo, muestra error
   - Los checkboxes tienen valores como: "Id", "TrnDate", "Amount", etc.

3. **Cierra el modal y ejecuta la descarga**:
   - Llama a `download()` con los parámetros

---

### 5. Método `download()` - EL MÁS IMPORTANTE

**Función**: Descarga datos del servidor OData e inserta en Excel

Este es el método más complejo y crítico del add-in.

**Flujo completo**:

```
download(downloadType, recordLimit, selectedFields)
    │
    ▼
┌─────────────────────────────────────────┐
│  1. PREPARACIÓN - Excel.run()          │
├─────────────────────────────────────────┤
│  • Suspender actualización de pantalla │
│    (optimización de rendimiento)        │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  2. CONSTRUCCIÓN DE URL OData          │
├─────────────────────────────────────────┤
│  Según downloadType:                    │
│                                         │
│  CUENTAS:                               │
│  └─► /odata/AccountSet?$top=50         │
│                                         │
│  FLUJOS:                                │
│  └─► /odata/FlowCodeSet?$top=50        │
│                                         │
│  MOVIMIENTOS:                           │
│  └─► /odata/CashFlowSet?               │
│      $select=Id,TrnDate,Amount&        │
│      $expand=FlowCode($select=Code),   │
│               Account(...),            │
│               TrnCurrency(...)&        │
│      $filter=Status eq 'Actual'&       │
│      $top=50                            │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  3. PETICIÓN AL SERVIDOR               │
├─────────────────────────────────────────┤
│  • Usa authenticatedFetch()            │
│  • Incluye header Authorization        │
│  • Sistema de reintentos (3 intentos) │
│  • Espera 1s entre intentos            │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  4. PROCESAMIENTO DE RESPUESTA         │
├─────────────────────────────────────────┤
│  const data = await response.json()    │
│  const records = data.value            │
│                                         │
│  Validación:                            │
│  • Verificar que records existe        │
│  • Verificar que no está vacío         │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  5. GESTIÓN DE HOJAS EN EXCEL          │
├─────────────────────────────────────────┤
│  Determinar nombre de hoja:            │
│  • Cuentas → "Accounts"                │
│  • Flujos → "Flujos"                   │
│  • Movimientos → "Movimientos"         │
│                                         │
│  Eliminar hoja si ya existe:           │
│  • worksheets.getItem(sheetName)       │
│  • existingSheet.delete()              │
│                                         │
│  Crear nueva hoja:                     │
│  • worksheets.add(sheetName)           │
│  • sheet.load(["protection", "name"])  │
│  • await context.sync()                │
│                                         │
│  Eliminar Sheet1 (primera vez):        │
│  • Solo si existe                      │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  6. FORMATEO DE DATOS                  │
├─────────────────────────────────────────┤
│  Función: formatValue(fieldName, value)│
│                                         │
│  • FECHAS:                              │
│    - Detecta campos tipo fecha         │
│    - Convierte a número serial Excel   │
│    - Fórmula: (ms - epoch) / msPerDay  │
│                                         │
│  • BOOLEANOS:                           │
│    - true → "true"                     │
│    - false → "false"                   │
│                                         │
│  • ID (números grandes):                │
│    - Agrega apóstrofe al inicio        │
│    - Fuerza formato texto              │
│    - Evita notación científica         │
│                                         │
│  • VALORES NULOS:                       │
│    - null/undefined → ""               │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  7. PREPARACIÓN DE ENCABEZADOS Y DATOS │
├─────────────────────────────────────────┤
│  Si hay campos seleccionados:          │
│  • headers = selectedFields            │
│  • values = solo esos campos           │
│                                         │
│  Si no:                                 │
│  • headers = todos los campos          │
│  • values = todos los valores          │
│                                         │
│  Excluir: @odata.etag                  │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  8. ESCRITURA EN EXCEL                 │
├─────────────────────────────────────────┤
│  Calcular rango:                        │
│  • numRows = records.length + 1        │
│  • numCols = headers.length            │
│  • endColumn = getColumnLetter()       │
│                                         │
│  Escribir en un solo bloque:           │
│  • range = sheet.getRange(             │
│      `A1:${endColumn}${numRows}`       │
│    )                                    │
│  • range.values = [headers, ...values] │
│                                         │
│  ¡UN SOLO SYNC PARA RENDIMIENTO!      │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  9. APLICAR FORMATO                    │
├─────────────────────────────────────────┤
│  Encabezados:                           │
│  • Fondo azul (#4472C4)                │
│  • Texto blanco (#FFFFFF)              │
│  • Negrita                              │
│                                         │
│  Columnas de fecha:                     │
│  • Formato: "DD/MM/YYYY"               │
│  • Aplicado a columnas específicas     │
│                                         │
│  Columna Id:                            │
│  • Formato: "@" (texto)                │
│  • Evita notación científica           │
│                                         │
│  Autoajustar columnas:                  │
│  • range.format.autofitColumns()       │
└─────────────────────────────────────────┘
    │
    ▼
┌─────────────────────────────────────────┐
│  10. FINALIZACIÓN                      │
├─────────────────────────────────────────┤
│  • sheet.activate()                     │
│  • await context.sync()                │
│  • showNotification("Éxito", "success")│
└─────────────────────────────────────────┘
```

**Código clave - Construcción de URL OData**:

```javascript
// Para MOVIMIENTOS con parámetros complejos
let endpoint = '/odata/CashFlowSet';
const params = [];

// Límite de registros
if (recordLimit !== 'all') {
  params.push(`$top=${recordLimit}`);
}

// Seleccionar solo campos específicos
if (selectedFields.length > 0) {
  params.push(`$select=${selectedFields.join(',')}`);
}

// Expandir entidades relacionadas
const expandParam = '$expand=FlowCode($select=Code),BudgetCode($select=Code),' +
                    'Account($expand=Master($select=Code);$select=Id),' +
                    'TrnCurrency($select=Id)';
params.push(expandParam);

// Filtrar solo registros con Status='Actual'
params.push("$filter=Status eq 'Actual'");

// Unir parámetros
if (params.length > 0) {
  endpoint += '?' + params.join('&');
}
```

**Resultado de URL**:
```
/odata/CashFlowSet?
  $top=50&
  $select=Id,TrnDate,Amount&
  $expand=FlowCode($select=Code),Account($expand=Master($select=Code);$select=Id)&
  $filter=Status eq 'Actual'
```

**Código clave - Formateo de fechas**:

```javascript
const formatDate = (dateString, fieldName) => {
  if (!dateString || dateString === '') return '';
  
  const date = new Date(dateString);
  
  // Verificar validez
  if (isNaN(date.getTime())) return '';
  
  // Convertir a número serial de Excel
  // Excel cuenta días desde 30/12/1899
  const excelEpoch = new Date(1899, 11, 30);
  const msPerDay = 24 * 60 * 60 * 1000;
  const excelSerialDate = (date.getTime() - excelEpoch.getTime()) / msPerDay;
  
  return excelSerialDate;
};
```

**Ejemplo de conversión**:
- Fecha: `2025-11-12T10:30:00Z`
- JavaScript: `new Date("2025-11-12T10:30:00Z")`
- Excel Serial: `46051.4375` (días desde 30/12/1899)
- Formato aplicado: `DD/MM/YYYY` → `12/11/2025`

**Código clave - Escritura eficiente en Excel**:

```javascript
// ❌ FORMA INEFICIENTE (múltiples syncs):
for (let i = 0; i < records.length; i++) {
  for (let j = 0; j < headers.length; j++) {
    sheet.getCell(i, j).values = [[records[i][headers[j]]]];
    await context.sync(); // SYNC POR CADA CELDA = MUY LENTO
  }
}

// ✅ FORMA EFICIENTE (un solo sync):
const range = sheet.getRange(`A1:${endColumn}${numRows}`);
range.values = [headers, ...values]; // ESCRIBIR TODO DE UNA VEZ
await context.sync(); // UN SOLO SYNC = RÁPIDO
```

**Optimización de rendimiento**:
- `suspendScreenUpdatingUntilNextSync()`: No actualiza la pantalla hasta el sync final
- Escritura en bloques: Un solo `range.values` con todos los datos
- Formatos en batch: Un `numberFormat` por columna completa, no por celda

---

## 🔧 Métodos y Funciones Principales

### 6. Función `authenticatedFetch()`

**Función**: Wrapper para hacer peticiones autenticadas al servidor

```javascript
async function authenticatedFetch(url, options = {}) {
  if (!userCredentials.isLoggedIn) {
    throw new Error("Debe iniciar sesión primero");
  }

  const defaultOptions = {
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      'Accept': 'application/json; charset=utf-8',
      'Authorization': `Basic ${btoa(userCredentials.username + ':' + userCredentials.password)}`
    }
  };

  return fetch(url, { ...defaultOptions, ...options });
}
```

**¿Qué hace?**

1. **Verificación de autenticación**:
   - Lanza error si no hay sesión iniciada

2. **Configuración automática de headers**:
   - Content-Type: JSON con UTF-8
   - Accept: JSON con UTF-8
   - Authorization: Basic [Base64]

3. **Merge de opciones**:
   - `{ ...defaultOptions, ...options }`
   - Permite sobreescribir opciones por defecto

**Uso**:
```javascript
const response = await authenticatedFetch('/odata/AccountSet?$top=10');
const data = await response.json();
```

---

### 7. Función `showNotification()`

**Función**: Muestra notificaciones temporales al usuario

```javascript
function showNotification(message, type = 'success') {
  const popup = document.getElementById('notificationPopup');
  const messageEl = document.getElementById('notificationMessage');
  
  // Establecer el mensaje
  messageEl.textContent = message;
  
  // Aplicar clase de estilo según el tipo
  popup.classList.remove('success', 'error');
  popup.classList.add(type);
  
  // Mostrar el popup
  popup.classList.remove('hidden');
  
  // Ocultar después de 3 segundos
  setTimeout(() => {
    popup.classList.add('hidden');
  }, 3000);
}
```

**Tipos de notificación**:
- `success`: Verde (#107C10) - Operación exitosa
- `error`: Rojo (#A4262C) - Error o advertencia

**Uso**:
```javascript
showNotification("¡Datos descargados exitosamente!", "success");
showNotification("Error: No se pudo conectar al servidor", "error");
```

---

### 8. Función `getColumnLetter()`

**Función**: Convierte índice numérico a letra de columna Excel

```javascript
const getColumnLetter = (colIndex) => {
  let letter = '';
  while (colIndex >= 0) {
    letter = String.fromCharCode((colIndex % 26) + 65) + letter;
    colIndex = Math.floor(colIndex / 26) - 1;
  }
  return letter;
};
```

**¿Cómo funciona?**

Sistema de base 26 (A-Z):
- 0 → A
- 25 → Z
- 26 → AA
- 27 → AB
- 701 → ZZ
- 702 → AAA

**Ejemplos**:
```javascript
getColumnLetter(0);   // "A"
getColumnLetter(1);   // "B"
getColumnLetter(25);  // "Z"
getColumnLetter(26);  // "AA"
getColumnLetter(27);  // "AB"
```

**Uso en el código**:
```javascript
const numCols = headers.length; // Ej: 15 columnas
const endColumn = getColumnLetter(numCols - 1); // "O"
const range = sheet.getRange(`A1:${endColumn}${numRows}`); // "A1:O51"
```

---

### 9. Método `importData()` (Función de ejemplo)

**Función**: Ejemplo de cómo enviar datos desde Excel a un servidor

```javascript
export async function importData() {
  try {
    await Excel.run(async (context) => {
      // 1. Obtener datos de Excel
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:B2");
      range.load(["values"]);
      await context.sync();

      // 2. Preparar datos para enviar
      const data = {
        title: range.values[1][0] || "",
        body: range.values[1][1] || "",
        userId: 1
      };

      // 3. Enviar a servidor de prueba
      const response = await fetch('https://jsonplaceholder.typicode.com/posts', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      const result = await response.json();

      // 4. Crear hoja con resultado
      const resultSheet = context.workbook.worksheets.add("Resultado");
      const resultRange = resultSheet.getRange("A1:C2");
      resultRange.values = [
        ["ID", "Estado", "Fecha"],
        [result.id, "Importado exitosamente", new Date().toLocaleString()]
      ];

      await context.sync();
      showNotification("¡Datos importados exitosamente!", "success");
    });
  } catch (error) {
    console.error("Error:", error);
    showNotification("Error al importar los datos", "error");
  }
}
```

**Nota**: Esta función es un ejemplo que usa un servidor de prueba (jsonplaceholder). No se usa en producción, pero muestra cómo enviar datos desde Excel a un servidor.

---

## 📁 Estructura de Archivos

```
ExcelRestAdding/
│
├── src/                          # Código fuente
│   ├── taskpane/
│   │   ├── taskpane.html        # Interfaz de usuario principal
│   │   ├── taskpane.js          # Lógica de negocio (690 líneas)
│   │   └── taskpane.css         # Estilos CSS
│   │
│   └── commands/
│       ├── commands.html         # Página para comandos de cinta
│       └── commands.js           # Lógica de comandos
│
├── assets/                       # Recursos estáticos
│   ├── icon-16.png              # Icono 16x16
│   ├── icon-32.png              # Icono 32x32
│   ├── icon-64.png              # Icono 64x64
│   ├── icon-80.png              # Icono 80x80
│   ├── icon-128.png             # Icono 128x128
│   └── logo-filled.png          # Logo completo
│
├── dist/                         # Build de producción (generado)
│   ├── taskpane.html
│   ├── taskpane.js              # JavaScript empaquetado y minificado
│   ├── polyfill.js              # Polyfills para compatibilidad
│   ├── commands.html
│   ├── commands.js
│   ├── manifest.xml             # Manifest de producción
│   ├── assets/                   # Iconos copiados
│   ├── privacy-policy.html      # Política de privacidad
│   ├── terms-of-use.html        # Términos de uso
│   └── support.html             # Página de soporte
│
├── sharepoint-distribution/      # Paquete para SharePoint
│   ├── manifest-production.xml
│   ├── icon-128.png
│   ├── icon-64.png
│   ├── icon-32.png
│   └── INSTRUCCIONES-INSTALACION.md
│
├── manifest.xml                  # Manifest de desarrollo (localhost:3000)
├── manifest-production.xml       # Manifest de producción (GitHub Pages)
│
├── webpack.config.js             # Configuración de Webpack
├── babel.config.json             # Configuración de Babel
├── package.json                  # Dependencias y scripts npm
│
├── privacy-policy.html           # Política de privacidad (raíz)
├── terms-of-use.html            # Términos de uso (raíz)
├── support.html                  # Soporte técnico (raíz)
├── Documentacion-AddIn-Azurriga.html  # Documentación de usuario
│
└── README.md                     # Documentación del proyecto
```

---

## ⚙️ Configuración de Webpack

### webpack.config.js

**Función**: Configurar cómo se empaqueta y sirve la aplicación

**Configuración clave para desarrollo**:

```javascript
devServer: {
  port: 3000,
  https: true,
  headers: {
    "Access-Control-Allow-Origin": "*",
  },
  proxy: [
    {
      context: ['/odata'],
      target: 'http://8cf33ac.online-server.cloud:1031',
      changeOrigin: true,
      secure: false,
      logLevel: 'debug'
    }
  ]
}
```

**¿Qué hace el proxy?**

1. **Recibe peticiones a**: `https://localhost:3000/odata/...`
2. **Las redirige a**: `http://8cf33ac.online-server.cloud:1031/odata/...`
3. **Ventajas**:
   - Evita errores de CORS
   - Evita Mixed Content (HTTPS → HTTP)
   - El navegador ve todo como localhost:3000

**Flujo de una petición**:
```
Add-in (taskpane.js)
    │
    │ fetch('/odata/AccountSet')
    ▼
Webpack Dev Server (localhost:3000)
    │
    │ Proxy detecta /odata
    │ Redirige a http://8cf33ac.online-server.cloud:1031
    ▼
Servidor OData (8cf33ac.online-server.cloud:1031)
    │
    │ Responde con datos
    ▼
Webpack Dev Server
    │
    │ Devuelve respuesta
    ▼
Add-in (taskpane.js)
    │
    │ const data = await response.json()
    ▼
Excel (muestra datos)
```

**Plugins de Webpack**:

```javascript
plugins: [
  new HtmlWebpackPlugin({
    filename: "taskpane.html",
    template: "./src/taskpane/taskpane.html",
    chunks: ["polyfill", "taskpane"]
  }),
  
  new CopyWebpackPlugin({
    patterns: [
      { from: "assets/*", to: "assets/[name][ext]" },
      { from: "manifest*.xml", to: "[name][ext]" }
    ]
  }),
  
  new webpack.ProvidePlugin({
    Promise: ["es6-promise", "Promise"]
  })
]
```

**¿Qué hacen?**

1. **HtmlWebpackPlugin**:
   - Genera `taskpane.html` en `dist/`
   - Inyecta automáticamente `<script>` para polyfill.js y taskpane.js
   
2. **CopyWebpackPlugin**:
   - Copia assets/ a dist/assets/
   - Copia manifests a dist/

3. **ProvidePlugin**:
   - Proporciona `Promise` globalmente
   - Polyfill para navegadores antiguos

---

## 🔄 Flujo Completo de Datos

### Escenario: Usuario descarga 100 cuentas

```
1. Usuario abre Excel
   │
   └─► Excel carga el add-in desde GitHub Pages o localhost
       (manifest.xml especifica la URL)
       │
       └─► Se descarga taskpane.html, taskpane.js, taskpane.css
           │
           └─► Office.onReady() se ejecuta
               │
               └─► Se muestran los botones: Login, Download, Import
                   (Download e Import están deshabilitados)

2. Usuario hace click en "Iniciar sesión"
   │
   └─► login() abre modal
       │
       └─► Usuario introduce: username="admin", password="pass123"
           │
           └─► Click en "Iniciar sesión"
               │
               ├─► Se crea: authString = btoa("admin:pass123")
               │             = "YWRtaW46cGFzczEyMw=="
               │
               ├─► Se hace fetch('/odata/')
               │   con header: Authorization: Basic YWRtaW46cGFzczEyMw==
               │   │
               │   └─► Webpack proxy redirige a:
               │       http://8cf33ac.online-server.cloud:1031/odata/
               │       │
               │       └─► Servidor valida credenciales
               │           │
               │           └─► ✅ Respuesta 200 OK
               │               │
               │               └─► userCredentials = {
               │                     username: "admin",
               │                     password: "pass123",
               │                     isLoggedIn: true
               │                   }
               │
               └─► Botón cambia a: "¡Bienvenido admin!" (verde)
                   Botones Download e Import se activan

3. Usuario hace click en "Descargar"
   │
   └─► showDownloadModal() abre modal
       │
       └─► Usuario selecciona:
           • Tipo: "Cuentas"
           • Límite: "100"
           │
           └─► Click en "Descargar"
               │
               └─► executeDownload() recopila opciones
                   │
                   └─► Llama a: download("cuentas", "100", [])
                       │
                       ├─► Excel.run(async (context) => {
                       │     // Todo el proceso dentro de Excel
                       │
                       ├─► Construye URL:
                       │   /odata/AccountSet?$top=100
                       │
                       ├─► authenticatedFetch() hace petición
                       │   │
                       │   └─► Proxy redirige a servidor OData
                       │       │
                       │       └─► Servidor responde con JSON:
                       │           {
                       │             "value": [
                       │               { "Id": "001", "Code": "CAJA", ... },
                       │               { "Id": "002", "Code": "BANCO", ... },
                       │               ... (100 registros)
                       │             ]
                       │           }
                       │
                       ├─► Procesa respuesta:
                       │   const records = data.value; // 100 registros
                       │
                       ├─► Gestiona hojas:
                       │   • Si existe hoja "Accounts" → eliminarla
                       │   • Crear nueva hoja "Accounts"
                       │   • Eliminar "Sheet1" si existe
                       │
                       ├─► Formatea datos:
                       │   headers = ["Id", "Code", "Description", "Balance", ...]
                       │   values = [
                       │     ["'001", "CAJA", "Caja general", 1500.50, ...],
                       │     ["'002", "BANCO", "Banco Santander", 25000.00, ...],
                       │     ... (100 filas)
                       │   ]
                       │
                       ├─► Escribe en Excel (UN SOLO SYNC):
                       │   range = sheet.getRange("A1:J101") // 10 columnas x 101 filas
                       │   range.values = [headers, ...values]
                       │
                       ├─► Aplica formato:
                       │   • Encabezados: azul, blanco, negrita
                       │   • Columnas de fecha: DD/MM/YYYY
                       │   • Columna Id: formato texto
                       │   • Autoajustar columnas
                       │
                       ├─► sheet.activate()
                       │   await context.sync()
                       │
                       └─► }) // Fin de Excel.run
                           │
                           └─► showNotification(
                                 "¡100 cuentas descargados exitosamente!",
                                 "success"
                               )

4. Usuario ve en Excel:
   │
   └─► Hoja "Accounts" con:
       • 101 filas (1 encabezado + 100 datos)
       • Encabezados en azul con texto blanco
       • Fechas formateadas como DD/MM/YYYY
       • Columnas ajustadas automáticamente
       • Notificación verde de éxito (3 segundos)
```

---

## 🔍 Conceptos Clave de Office.js

### `Office.run()` vs `Excel.run()`

**Office.run()**: Para operaciones generales de Office
**Excel.run()**: Para operaciones específicas de Excel

```javascript
await Excel.run(async (context) => {
  // context es el contexto de ejecución de Excel
  // Todas las operaciones de Excel usan este context
  
  const sheet = context.workbook.worksheets.getActiveWorksheet();
  // sheet es un objeto PROXY, no tiene datos aún
  
  sheet.load("name"); // Decimos qué propiedades necesitamos
  
  await context.sync(); // AQUÍ se ejecutan todas las operaciones
  // Después del sync, sheet.name ya tiene valor
  
  console.log(sheet.name); // "Hoja1"
});
```

### Load y Sync

**load()**: Marca qué propiedades necesitas leer
**sync()**: Ejecuta todas las operaciones pendientes

```javascript
// ❌ MAL - No funciona
const sheet = context.workbook.worksheets.getActiveWorksheet();
console.log(sheet.name); // undefined - no se ha cargado

// ✅ BIEN
const sheet = context.workbook.worksheets.getActiveWorksheet();
sheet.load("name");
await context.sync();
console.log(sheet.name); // "Hoja1"
```

### Operaciones en Batch

**Concepto**: Acumular operaciones y ejecutarlas todas de una vez

```javascript
// ❌ FORMA LENTA (múltiples syncs)
for (let i = 0; i < 100; i++) {
  range.getCell(i, 0).values = [[data[i]]];
  await context.sync(); // 100 syncs = LENTO
}

// ✅ FORMA RÁPIDA (un solo sync)
range.values = data; // Asignar todo de una vez
await context.sync(); // 1 sync = RÁPIDO
```

---

## 📊 Ejemplo de Datos OData

### Respuesta de `/odata/AccountSet?$top=3`

```json
{
  "@odata.context": "http://8cf33ac.online-server.cloud:1031/odata/$metadata#AccountSet",
  "value": [
    {
      "@odata.etag": "W/\"datetime'2025-11-12T10%3A30%3A00.0000000'\"",
      "Id": "57000000000001",
      "Code": "100.001",
      "Description": "Caja general",
      "Balance": 1500.50,
      "Active": true,
      "CreationDateTime": "2025-01-15T08:00:00Z",
      "ModificationDateTime": "2025-11-10T15:30:00Z"
    },
    {
      "@odata.etag": "W/\"datetime'2025-11-12T10%3A30%3A00.0000000'\"",
      "Id": "57000000000002",
      "Code": "100.002",
      "Description": "Banco Santander",
      "Balance": 25000.00,
      "Active": true,
      "CreationDateTime": "2025-01-15T08:00:00Z",
      "ModificationDateTime": "2025-11-12T09:15:00Z"
    },
    {
      "@odata.etag": "W/\"datetime'2025-11-12T10%3A30%3A00.0000000'\"",
      "Id": "57000000000003",
      "Code": "200.001",
      "Description": "Clientes varios",
      "Balance": 12500.75,
      "Active": true,
      "CreationDateTime": "2025-01-15T08:00:00Z",
      "ModificationDateTime": "2025-11-11T14:20:00Z"
    }
  ]
}
```

### Cómo se procesa

```javascript
// 1. Se recibe la respuesta
const data = await response.json();

// 2. Se extraen los registros
const records = data.value; // Array de 3 objetos

// 3. Se extraen los encabezados (excluyendo @odata.etag)
const headers = ["Id", "Code", "Description", "Balance", "Active", 
                "CreationDateTime", "ModificationDateTime"];

// 4. Se formatean los valores
const values = records.map(record => [
  "'57000000000001",           // Id como texto
  "100.001",                   // Code
  "Caja general",              // Description
  1500.50,                     // Balance
  "true",                      // Active (booleano → texto)
  45677.333333,                // CreationDateTime (fecha → serial Excel)
  46046.395833                 // ModificationDateTime (fecha → serial Excel)
]);

// 5. Se escribe en Excel
range.values = [headers, ...values];
// Resultado en Excel:
// A1: "Id"  B1: "Code"  C1: "Description"  ...
// A2: 57000000000001  B2: "100.001"  C2: "Caja general"  ...
```

---

## 🚀 Optimizaciones Implementadas

### 1. Suspensión de actualización de pantalla

```javascript
const application = context.workbook.application;
application.suspendScreenUpdatingUntilNextSync();
```

**Efecto**: Excel no redibuja la pantalla hasta que se complete todo el proceso.
**Beneficio**: Mejora significativa en velocidad (hasta 10x más rápido)

### 2. Escritura en bloques

```javascript
// En lugar de escribir celda por celda
range.values = [headers, ...values]; // Toda la matriz de una vez
```

**Beneficio**: Reduce de N syncs a 1 solo sync

### 3. Sistema de reintentos

```javascript
let retries = 3;
while (retries > 0) {
  try {
    response = await authenticatedFetch(endpoint);
    if (response.ok) break;
  } catch (fetchError) {
    retries--;
    if (retries === 0) throw new Error('Error después de 3 intentos');
    await new Promise(resolve => setTimeout(resolve, 1000)); // Esperar 1s
  }
}
```

**Beneficio**: Mayor resiliencia ante problemas de red temporales

### 4. Formateo eficiente

```javascript
// En lugar de formatear celda por celda
const dateRange = sheet.getRange(`${colLetter}2:${colLetter}${numRows}`);
dateRange.numberFormat = [["DD/MM/YYYY"]]; // Toda la columna de una vez
```

**Beneficio**: Una operación en lugar de N operaciones

---

## 🐛 Manejo de Errores

### Estrategia de manejo de errores

```javascript
try {
  await Excel.run(async (context) => {
    // Operaciones de Excel
  });
} catch (error) {
  console.error("Error específico:", error.message);
  
  let errorMessage = "Error al descargar los datos";
  
  // Mensajes específicos según el tipo de error
  if (error.message.includes("protegida")) {
    errorMessage = "La hoja está protegida. Desproteja la hoja e intente nuevamente.";
  } else if (error.message.includes("obtener datos")) {
    errorMessage = "Error de conexión. Verifique su conexión a internet.";
  }
  
  showNotification(errorMessage, "error");
}
```

### Tipos de errores manejados

1. **Errores de autenticación**:
   - 401 Unauthorized
   - 403 Forbidden
   - Credenciales incorrectas

2. **Errores de red**:
   - Failed to fetch
   - NetworkError
   - Timeout

3. **Errores de Excel**:
   - Hoja protegida
   - Permisos insuficientes
   - Rango inválido

4. **Errores de datos**:
   - Respuesta vacía
   - Formato incorrecto
   - Datos nulos

---

## 🔐 Seguridad

### Autenticación

- **Método**: HTTP Basic Authentication
- **Codificación**: Base64 (NO es encriptación)
- **Transmisión**: Siempre sobre HTTPS en producción
- **Almacenamiento**: Solo en memoria (userCredentials), nunca persistido

### Buenas prácticas implementadas

1. **No persistir credenciales**:
   ```javascript
   // Solo en memoria, se pierde al cerrar Excel
   let userCredentials = { username, password, isLoggedIn };
   ```

2. **HTTPS obligatorio en producción**:
   - GitHub Pages: HTTPS automático
   - Webpack Dev Server: HTTPS con certificados auto-firmados

3. **Validación de sesión**:
   ```javascript
   if (!userCredentials.isLoggedIn) {
     showNotification("Debe iniciar sesión primero", "error");
     return;
   }
   ```

4. **Headers de seguridad**:
   ```javascript
   headers: {
     'Content-Type': 'application/json; charset=utf-8',
     'Accept': 'application/json; charset=utf-8'
   }
   ```

---

## 📝 Resumen de Métodos por Responsabilidad

### Inicialización
- `Office.onReady()`: Punto de entrada cuando Office está listo

### Autenticación
- `login()`: Gestiona todo el flujo de login
- `authenticatedFetch()`: Wrapper para peticiones autenticadas

### Descarga de Datos
- `showDownloadModal()`: Muestra opciones de descarga
- `executeDownload()`: Recopila opciones y ejecuta
- `download()`: **MÉTODO PRINCIPAL** - Descarga e inserta datos en Excel
  - `formatValue()`: Formatea valores según tipo de campo
  - `formatDate()`: Convierte fechas a serial de Excel
  - `getColumnLetter()`: Convierte índice a letra de columna

### UI y Notificaciones
- `showNotification()`: Muestra mensajes temporales al usuario

### Ejemplo de Importación (No usado en producción)
- `importData()`: Ejemplo de cómo enviar datos desde Excel a un servidor

---

## 🎯 Próximos Pasos para Desarrollo

### Mejoras sugeridas

1. **Caché de datos**:
   ```javascript
   const dataCache = new Map();
   // Evitar descargar los mismos datos múltiples veces
   ```

2. **Paginación**:
   ```javascript
   // Para conjuntos de datos muy grandes
   let skip = 0;
   const top = 1000;
   while (hasMore) {
     await download(`${endpoint}?$skip=${skip}&$top=${top}`);
     skip += top;
   }
   ```

3. **Filtros personalizados**:
   - Permitir al usuario especificar filtros OData
   - Ejemplo: `$filter=Balance gt 1000 and Active eq true`

4. **Exportar a otros formatos**:
   - CSV
   - JSON
   - XML

5. **Modo offline**:
   - Service Workers para caché
   - IndexedDB para almacenamiento local

---

## 📚 Referencias

- **Office.js API**: https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview
- **OData v4**: https://www.odata.org/documentation/
- **Webpack**: https://webpack.js.org/
- **GitHub Pages**: https://pages.github.com/

---

**Documento creado**: 12/11/2025  
**Versión del Add-in**: 1.0.0.0  
**Autor**: Documentación técnica generada para Add-in Azurriga
