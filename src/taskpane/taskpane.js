/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/**
 * Variable global para almacenar las credenciales del usuario autenticado
 * @type {Object}
 * @property {string|null} username - Nombre de usuario
 * @property {string|null} password - Contraseña del usuario
 * @property {boolean} isLoggedIn - Indica si hay una sesión activa
 */
let userCredentials = {
  username: null,
  password: null,
  isLoggedIn: false
};

/**
 * Variable global para almacenar el ID de la cuenta seleccionada
 * @type {string|null}
 */
let selectedAccountId = null;

/**
 * Variable global para almacenar las cuentas descargadas desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
let cachedAccounts = [];

/**
 * Variable global para almacenar los flujos de caja descargados desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
let cachedFlowCodes = [];

/**
 * Variable global para almacenar los códigos presupuestarios descargados desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
let cachedBudgetCodes = [];

/**
 * Variable global para almacenar las divisas descargadas desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
let cachedCurrencies = [];

/**
 * Variable global para almacenar los lugares de cotización descargados desde OData
 * @type {Array<{Id: number, Description: string}>}
 */
let cachedQuotationPlaces = [];

/**
 * Variables globales para almacenar el rango de fechas seleccionado
 * @type {string|null}
 */
let selectedDateFrom = null;
let selectedDateTo = null;

/**
 * Función de inicialización de Office.js
 * Se ejecuta cuando el entorno de Office está listo para interactuar
 * @param {Object} info - Información sobre el host de Office
 */
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
        // Cargar cuentas al mostrar opciones de movimientos
        loadAccounts();
      } else {
        movimientosOptions.classList.add("hidden");
      }
    };
    
    // Event listener para cambio de cuenta seleccionada
    document.getElementById("accountSelect").onchange = function() {
      selectedAccountId = this.value;
      console.log("Cuenta seleccionada:", this.options[this.selectedIndex].text, "(ID:", selectedAccountId, ")");
    };
    
    // Event listeners para los campos de fecha
    document.getElementById("dateFrom").onchange = function() {
      selectedDateFrom = this.value;
      console.log("Fecha desde seleccionada:", selectedDateFrom);
    };
    
    document.getElementById("dateTo").onchange = function() {
      selectedDateTo = this.value;
      console.log("Fecha hasta seleccionada:", selectedDateTo);
    };
    
    // Event listener para cerrar el panel de errores detallados
    document.getElementById("closeErrorDetails").onclick = hideErrorDetails;
  }
});

/**
 * Gestiona el proceso completo de autenticación del usuario
 * 
 * Flujo:
 * 1. Muestra un modal de inicio de sesión
 * 2. Captura las credenciales (usuario y contraseña)
 * 3. Valida las credenciales contra el servidor OData usando HTTP Basic Auth
 * 4. Si la autenticación es exitosa:
 *    - Guarda las credenciales en memoria
 *    - Actualiza la UI (botón de login y habilita otras funciones)
 *    - Muestra notificación de éxito
 * 5. Si falla:
 *    - Muestra mensaje de error en el modal
 *    - Permite reintentar
 * 
 * @async
 * @throws {Error} Si hay problemas de conexión o el modal no se encuentra en el DOM
 */
export async function login() {
  try {
    console.log("Función login iniciada");
    const modal = document.getElementById("loginModal");
    if (!modal) {
      console.error("Modal no encontrado en el DOM");
      return;
    }
    console.log("Modal encontrado, removiendo clase hidden");
    modal.classList.remove("hidden");
    modal.style.display = "block"; // Forzar visualización

    const loginSubmitButton = document.getElementById("loginSubmit");
    if (!loginSubmitButton) {
      console.error("Botón submit no encontrado");
      return;
    }
    console.log("Configurando evento click del botón submit");
    loginSubmitButton.onclick = async () => {
      const username = document.getElementById("username").value;
      const password = document.getElementById("password").value;

      if (!username || !password) {
        console.error("Por favor complete todos los campos");
        return;
      }

      // Mostrar spinner y ocultar error previo
      const loadingDiv = document.getElementById("loginLoading");
      const errorDiv = document.getElementById("loginError");
      const submitButton = document.getElementById("loginSubmit");
      const cancelButton = document.getElementById("loginCancel");
      
      loadingDiv.classList.remove("hidden");
      errorDiv.classList.add("hidden");
      submitButton.disabled = true;
      cancelButton.disabled = true;

      try {
        console.log("Intentando hacer login con:", { username });
        
        // Crear el header de autenticación básica
        const authString = btoa(username + ':' + password);
        console.log("Autenticación básica creada");
        
        // DESARROLLO: Usar proxy de webpack (/odata)
        // PRODUCCIÓN: Usar proxy Vercel (https://excel-addin-azurriga.vercel.app)
        const isDevelopment = window.location.hostname === 'localhost';
        const baseUrl = isDevelopment 
          ? '/odata/AccountSet?$top=1'
          : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/AccountSet?$top=1';
        
        console.log(`Usando proxy ${isDevelopment ? 'WEBPACK' : 'VERCEL'}: ${baseUrl}`);
        
        const response = await fetch(baseUrl, {
          method: 'GET',
          headers: {
            'Authorization': `Basic ${authString}`,
            'Content-Type': 'application/json',
          }
        });

        console.log("Respuesta recibida:", response);
        console.log("Status:", response.status);
        console.log("Status Text:", response.statusText);

        // Ocultar spinner
        loadingDiv.classList.add("hidden");
        submitButton.disabled = false;
        cancelButton.disabled = false;

        if (response.ok) {
          console.log("Login exitoso");
          
          // Guardar las credenciales
          userCredentials.username = username;
          userCredentials.password = password;
          userCredentials.isLoggedIn = true;
          
          const loginButton = document.getElementById("login");
          loginButton.innerHTML = `<span class="ms-Button-label">¡Bienvenido ${username}!</span>`;
          loginButton.style.backgroundColor = "#107C10";
          
          // Activar los botones de Descargar e Importar
          const downloadButton = document.getElementById("download");
          const importButton = document.getElementById("import");
          downloadButton.classList.remove("is-disabled");
          downloadButton.removeAttribute("disabled");
          importButton.classList.remove("is-disabled");
          importButton.removeAttribute("disabled");
          
          modal.classList.add("hidden");
          
          showNotification("¡Sesión iniciada correctamente!", "success");
        } else {
          console.error("Error de autenticación. Status:", response.status);
          
          // Leer el cuerpo de la respuesta para más detalles
          let errorDetails = '';
          try {
            const errorText = await response.text();
            errorDetails = ` (${response.status}: ${errorText.substring(0, 100)})`;
          } catch (e) {
            errorDetails = ` (Código: ${response.status})`;
          }
          
          // Mostrar mensaje de error en el modal
          const errorDiv = document.getElementById("loginError");
          errorDiv.innerHTML = `Usuario o contraseña incorrectos${errorDetails}`;
          errorDiv.classList.remove("hidden");
          
          // Limpiar el mensaje de error después de 5 segundos
          setTimeout(() => {
            errorDiv.classList.add("hidden");
          }, 5000);
        }
      } catch (error) {
        console.error("Error en login (catch):", error);
        console.error("Error message:", error.message);
        console.error("Error stack:", error.stack);
        
        // Ocultar spinner y reactivar botones
        loadingDiv.classList.add("hidden");
        submitButton.disabled = false;
        cancelButton.disabled = false;
        
        const errorDiv = document.getElementById("loginError");
        
        // Construir mensaje de error más detallado
        let errorMsg = "Error de conexión: ";
        if (error.message.includes('Failed to fetch')) {
          errorMsg += "No se puede conectar al servidor. Verifique:\n1. La URL del servidor\n2. Que el servidor esté en ejecución\n3. Configuración de CORS en el servidor";
        } else if (error.message.includes('NetworkError')) {
          errorMsg += "Error de red. Verifique su conexión a Internet.";
        } else {
          errorMsg += error.message;
        }
        
        errorDiv.innerHTML = errorMsg.replace(/\n/g, '<br>');
        errorDiv.classList.remove("hidden");
        
        // Limpiar el mensaje de error después de 7 segundos
        setTimeout(() => {
          errorDiv.classList.add("hidden");
        }, 7000);
      }
    };

    document.getElementById("loginCancel").onclick = () => {
      modal.classList.add("hidden");
    };

    window.onclick = (event) => {
      if (event.target === modal) {
        modal.classList.add("hidden");
      }
    };
  } catch (error) {
    console.error("Error:", error);
  }
}

/**
 * Muestra una notificación temporal (popup) al usuario
 * 
 * @param {string} message - Mensaje a mostrar
 * @param {string} [type='success'] - Tipo de notificación: 'success' (verde) o 'error' (rojo)
 * 
 * Características:
 * - Se muestra durante 3 segundos
 * - Se aplican estilos diferentes según el tipo
 * - Se auto-oculta automáticamente
 */
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

/**
 * Muestra un panel modal con errores detallados
 * @param {string} message - Mensaje detallado de errores
 */
function showErrorDetails(message) {
  const panel = document.getElementById('errorDetailsPanel');
  const messageEl = document.getElementById('errorDetailsMessage');
  
  // Establecer el mensaje
  messageEl.textContent = message;
  
  // Mostrar el panel
  panel.classList.remove('hidden');
}

/**
 * Oculta el panel de errores detallados
 */
function hideErrorDetails() {
  const panel = document.getElementById('errorDetailsPanel');
  panel.classList.add('hidden');
}

/**
 * Carga las cuentas activas desde el servidor y las muestra en el combo
 * 
 * Consulta: odata/AccountSet?$filter=Active eq true&$select=Code,Id
 * - Muestra el Code en el combo
 * - Almacena el Id al seleccionar una cuenta
 * 
 * @async
 */
async function loadAccounts() {
  try {
    const accountSelect = document.getElementById("accountSelect");
    
    // Limpiar opciones existentes (excepto la primera "Todas las cuentas")
    accountSelect.innerHTML = '<option value="">Todas las cuentas</option>';
    
    // Mostrar indicador de carga
    const loadingOption = document.createElement('option');
    loadingOption.value = '';
    loadingOption.textContent = 'Cargando cuentas...';
    loadingOption.disabled = true;
    accountSelect.appendChild(loadingOption);
    
    // Determinar el proxy correcto
    const isDevelopment = window.location.hostname === 'localhost';
    const VERCEL_PROXY = isDevelopment
      ? '/odata/'
      : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
    
    // Construir URL con filtro y select
    const separator = isDevelopment ? '?' : '%3F';
    const ampersand = isDevelopment ? '&' : '%26';
    const endpoint = `${VERCEL_PROXY}AccountSet${separator}$filter=Active eq true${ampersand}$select=Code,Id`;
    
    console.log("Cargando cuentas desde:", endpoint);
    
    // Obtener las cuentas con header especial para números grandes
    const response = await authenticatedFetch(endpoint, {
      headers: {
        'Accept': 'application/json;IEEE754Compatible=true'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Error al cargar cuentas: ${response.status}`);
    }
    
    const data = await response.json();
    console.log("Cuentas recibidas:", data);
    
    // Limpiar el indicador de carga
    accountSelect.innerHTML = '<option value="">Todas las cuentas</option>';
    
    // Verificar que tengamos datos
    if (data && data.value && data.value.length > 0) {
      // Almacenar las cuentas en el caché global
      cachedAccounts = data.value;
      
      // Agregar cada cuenta al combo
      data.value.forEach(account => {
        const option = document.createElement('option');
        option.value = account.Id;  // Valor interno: ID
        option.textContent = account.Code;  // Texto visible: Code
        accountSelect.appendChild(option);
      });
      
      console.log(`${data.value.length} cuentas cargadas correctamente`);
    } else {
      // No hay cuentas activas
      const noDataOption = document.createElement('option');
      noDataOption.value = '';
      noDataOption.textContent = 'No hay cuentas activas disponibles';
      noDataOption.disabled = true;
      accountSelect.appendChild(noDataOption);
    }
    
  } catch (error) {
    console.error("Error cargando cuentas:", error);
    
    // Mostrar error en el combo
    const accountSelect = document.getElementById("accountSelect");
    accountSelect.innerHTML = '<option value="">Error al cargar cuentas</option>';
    
    // Mostrar notificación al usuario
    showNotification("Error al cargar las cuentas: " + error.message, "error");
  }
}

/**
 * Carga los flujos de caja desde el servidor
 * 
 * Consulta: odata/FlowCodeSet?$select=Code,Id
 * 
 * @async
 */
async function loadFlowCodes() {
  try {
    console.log("Iniciando carga de flujos de caja...");
    
    const isDevelopment = window.location.hostname === 'localhost';
    const VERCEL_PROXY = isDevelopment
      ? '/odata/'
      : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
    
    const separator = isDevelopment ? '?' : '%3F';
    const endpoint = `${VERCEL_PROXY}FlowCodeSet${separator}$select=Code,Id`;
    
    console.log("Cargando flujos desde:", endpoint);
    
    const response = await authenticatedFetch(endpoint, {
      headers: {
        'Accept': 'application/json;IEEE754Compatible=true'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Error al cargar flujos: ${response.status}`);
    }
    
    const data = await response.json();
    console.log("Flujos recibidos:", data);
    
    if (data && data.value && data.value.length > 0) {
      cachedFlowCodes = data.value;
      console.log(`${data.value.length} flujos de caja cargados correctamente`);
    } else {
      console.warn("No hay flujos de caja disponibles");
    }
  } catch (error) {
    console.error("Error cargando flujos de caja:", error);
    showNotification("Error al cargar flujos de caja: " + error.message, "error");
  }
}

/**
 * Carga los códigos presupuestarios desde el servidor
 * 
 * Consulta: odata/BudgetCodeSet?$select=Code,Id
 * 
 * @async
 */
async function loadBudgetCodes() {
  try {
    console.log("Iniciando carga de códigos presupuestarios...");
    
    const isDevelopment = window.location.hostname === 'localhost';
    const VERCEL_PROXY = isDevelopment
      ? '/odata/'
      : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
    
    const separator = isDevelopment ? '?' : '%3F';
    const endpoint = `${VERCEL_PROXY}BudgetCodeSet${separator}$select=Code,Id`;
    
    console.log("Cargando códigos presupuestarios desde:", endpoint);
    
    const response = await authenticatedFetch(endpoint, {
      headers: {
        'Accept': 'application/json;IEEE754Compatible=true'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Error al cargar códigos presupuestarios: ${response.status}`);
    }
    
    const data = await response.json();
    console.log("Códigos presupuestarios recibidos:", data);
    
    if (data && data.value && data.value.length > 0) {
      cachedBudgetCodes = data.value;
      console.log(`${data.value.length} códigos presupuestarios cargados correctamente`);
    } else {
      console.warn("No hay códigos presupuestarios disponibles");
    }
  } catch (error) {
    console.error("Error cargando códigos presupuestarios:", error);
    showNotification("Error al cargar códigos presupuestarios: " + error.message, "error");
  }
}

/**
 * Carga las divisas desde el servidor
 * 
 * Consulta: odata/CurrencySet (sin $select para obtener todos los campos)
 * 
 * @async
 */
async function loadCurrencies() {
  try {
    console.log("Iniciando carga de divisas...");
    
    const isDevelopment = window.location.hostname === 'localhost';
    const VERCEL_PROXY = isDevelopment
      ? '/odata/'
      : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
    
    // No usar $select para obtener todos los campos disponibles
    const endpoint = `${VERCEL_PROXY}CurrencySet`;
    
    console.log("Cargando divisas desde:", endpoint);
    
    const response = await authenticatedFetch(endpoint, {
      headers: {
        'Accept': 'application/json;IEEE754Compatible=true'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Error al cargar divisas: ${response.status}`);
    }
    
    const data = await response.json();
    console.log("Divisas recibidas:", data);
    
    if (data && data.value && data.value.length > 0) {
      // Mapear los datos según la estructura real de CurrencySet
      // Asumiendo que tiene Id como primary key
      cachedCurrencies = data.value.map(curr => ({
        Id: curr.Id || curr.Code || curr.id,
        Code: curr.Code || curr.Id || curr.id
      }));
      console.log(`${cachedCurrencies.length} divisas cargadas correctamente`);
      console.log("Ejemplo de divisa:", cachedCurrencies[0]);
    } else {
      console.warn("No hay divisas disponibles");
    }
  } catch (error) {
    console.error("Error cargando divisas:", error);
    showNotification("Error al cargar divisas: " + error.message, "error");
  }
}

/**
 * Carga los lugares de cotización desde el servidor
 * 
 * Consulta: odata/QuotationPlaceSet (sin $select para obtener todos los campos)
 * 
 * @async
 */
async function loadQuotationPlaces() {
  try {
    console.log("Iniciando carga de lugares de cotización...");
    
    const isDevelopment = window.location.hostname === 'localhost';
    const VERCEL_PROXY = isDevelopment
      ? '/odata/'
      : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
    
    // No usar $select para obtener todos los campos disponibles
    const endpoint = `${VERCEL_PROXY}QuotationPlaceSet`;
    
    console.log("Cargando lugares de cotización desde:", endpoint);
    
    const response = await authenticatedFetch(endpoint, {
      headers: {
        'Accept': 'application/json;IEEE754Compatible=true'
      }
    });
    
    if (!response.ok) {
      throw new Error(`Error al cargar lugares de cotización: ${response.status}`);
    }
    
    const data = await response.json();
    console.log("Lugares de cotización recibidos:", data);
    
    if (data && data.value && data.value.length > 0) {
      // Mapear los datos según la estructura real de QuotationPlaceSet
      cachedQuotationPlaces = data.value.map(qp => ({
        Id: qp.Id || qp.id,
        Description: qp.Description || qp.Name || qp.description || qp.name || `Cotización ${qp.Id}`
      }));
      console.log(`${cachedQuotationPlaces.length} lugares de cotización cargados correctamente`);
      console.log("Ejemplo de lugar de cotización:", cachedQuotationPlaces[0]);
    } else {
      console.warn("No hay lugares de cotización disponibles");
    }
  } catch (error) {
    console.error("Error cargando lugares de cotización:", error);
    showNotification("Error al cargar lugares de cotización: " + error.message, "error");
  }
}

/**
 * Wrapper para realizar peticiones HTTP autenticadas al servidor OData
 * 
 * Agrega automáticamente:
 * - Header de autenticación HTTP Basic (Base64)
 * - Headers de Content-Type y Accept para JSON
 * 
 * @async
 * @param {string} url - URL del endpoint a consultar
 * @param {Object} [options={}] - Opciones adicionales de fetch (se fusionan con las opciones por defecto)
 * @returns {Promise<Response>} Promesa con la respuesta HTTP
 * @throws {Error} Si el usuario no ha iniciado sesión
 */
async function authenticatedFetch(url, options = {}) {
  if (!userCredentials.isLoggedIn) {
    throw new Error("Debe iniciar sesión primero");
  }

  const defaultHeaders = {
    'Content-Type': 'application/json; charset=utf-8',
    'Accept': 'application/json; charset=utf-8',
    'Authorization': `Basic ${btoa(userCredentials.username + ':' + userCredentials.password)}`
  };

  // Mezclar headers personalizados con los predeterminados
  const mergedOptions = {
    ...options,
    headers: {
      ...defaultHeaders,
      ...(options.headers || {})
    }
  };

  return fetch(url, mergedOptions);
}

/**
 * Muestra el modal de configuración de descarga
 * 
 * Permite al usuario configurar:
 * - Tipo de descarga: Cuentas, Flujos de caja, Códigos Presupuestarios, Divisas, Cotización o Movimientos
 * - Límite de registros: 50, 100, 500 o todos
 * - Campos específicos (solo para Movimientos)
 * 
 * @async
 * @throws {Error} Si el usuario no está autenticado o hay problemas con el DOM
 */
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

    // Cerrar modal al hacer clic fuera
    window.onclick = (event) => {
      if (event.target === modal) {
        modal.classList.add("hidden");
      }
    };
  } catch (error) {
    console.error("Error al abrir modal de descarga:", error);
    showNotification("Error al abrir el modal de descarga", "error");
  }
}

/**
 * Recopila las opciones seleccionadas del modal y ejecuta la descarga
 * 
 * Validaciones:
 * - Para Movimientos: verifica que se haya seleccionado al menos un campo
 * 
 * @async
 * @throws {Error} Si no se cumplen las validaciones o hay problemas al preparar la descarga
 */
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

    console.log("Descarga:", downloadType, "| Registros:", recordLimit, "| Cuenta:", selectedAccountId || "Todas", "| Desde:", selectedDateFrom || "N/A", "| Hasta:", selectedDateTo || "N/A");

    // Cerrar el modal
    document.getElementById("downloadModal").classList.add("hidden");

    // Llamar a la función de descarga con los parámetros
    await download(downloadType, recordLimit, selectedFields);
  } catch (error) {
    console.error("Error en executeDownload:", error);
    showNotification("Error al preparar la descarga", "error");
  }
}

/**
 * ⭐ MÉTODO PRINCIPAL - Descarga datos del servidor OData e inserta en Excel
 * 
 * Proceso completo:
 * 1. Suspende actualización de pantalla (optimización)
 * 2. Construye URL OData con parámetros ($top, $select, $expand, $filter)
 * 3. Realiza petición autenticada al servidor (con reintentos)
 * 4. Procesa la respuesta JSON
 * 5. Gestiona hojas en Excel (elimina si existe, crea nueva)
 * 6. Formatea datos (fechas a serial Excel, booleanos, IDs como texto)
 * 7. Escribe datos en Excel en UN SOLO BLOQUE (optimización)
 * 8. Aplica formato visual (encabezados azules, formato de fecha, autoajuste)
 * 9. Activa la hoja y muestra notificación de éxito
 * 
 * @async
 * @param {string} [downloadType='cuentas'] - Tipo de datos: 'cuentas', 'flujos', 'codigos-presupuestarios', 'divisas', 'cotizacion' o 'movimientos'
 * @param {string} [recordLimit='50'] - Límite de registros: '50', '100', '500' o 'all'
 * @param {Array<string>} [selectedFields=[]] - Campos seleccionados (solo para movimientos)
 * 
 * Optimizaciones implementadas:
 * - suspendScreenUpdatingUntilNextSync() - Evita redibujado durante la operación
 * - Escritura en bloques - Un solo sync en lugar de N syncs
 * - Sistema de reintentos - 3 intentos en caso de fallo de red
 * 
 * @throws {Error} Si hay problemas de conexión, hoja protegida o formato de datos incorrecto
 */
export async function download(downloadType = 'cuentas', recordLimit = '50', selectedFields = []) {
  try {
    // Suspender actualización de pantalla para mejor rendimiento
    await Excel.run(async (context) => {
      const application = context.workbook.application;
      application.suspendScreenUpdatingUntilNextSync();
      
      // DESARROLLO: Usar proxy de webpack (/odata)
      // PRODUCCIÓN: Usar proxy Vercel (https://excel-addin-azurriga.vercel.app)
      const isDevelopment = window.location.hostname === 'localhost';
      const VERCEL_PROXY = isDevelopment
        ? '/odata/'
        : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
      
      console.log(`Download usando proxy ${isDevelopment ? 'WEBPACK' : 'VERCEL'}`);
      
      let endpoint = '';
      switch(downloadType) {
        case 'cuentas':
          endpoint = `${VERCEL_PROXY}AccountSet`;
          break;
        case 'flujos':
          endpoint = `${VERCEL_PROXY}FlowCodeSet`;
          break;
        case 'codigos-presupuestarios':
          endpoint = `${VERCEL_PROXY}BudgetCodeSet`;
          // Especificar los campos solicitados: Code, Id, Description
          const budgetParams = ['$select=Code,Id,Description'];
          if (recordLimit !== 'all') {
            budgetParams.push(`$top=${recordLimit}`);
          }
          if (budgetParams.length > 0) {
            const questionMark = isDevelopment ? '?' : '%3F';
            const ampersand = isDevelopment ? '&' : '%26';
            endpoint += questionMark + budgetParams.join(ampersand);
          }
          break;
        case 'divisas':
          endpoint = `${VERCEL_PROXY}CurrencySet`;
          break;
        case 'cotizacion':
          endpoint = `${VERCEL_PROXY}QuotationPlaceSet`;
          break;
        case 'movimientos':
          endpoint = `${VERCEL_PROXY}CashFlowSet`;
          // Construir la URL completa con $select, $expand y $filter
          const params = [];
          
          // Agregar límite de registros si no es "all"
          if (recordLimit !== 'all') {
            params.push(`$top=${recordLimit}`);
          }
          
          // Agregar $select con los campos seleccionados
          if (selectedFields.length > 0) {
            params.push(`$select=${selectedFields.join(',')}`);
          }
          
          // Agregar $expand (siempre se incluye para Movimientos)
          const expandParam = '$expand=FlowCode($select=Code),BudgetCode($select=Code),Account($expand=Master($select=Code);$select=Id),TrnCurrency($select=Id)';
          params.push(expandParam);
          
          // Construir $filter con Status y opcionalmente con AccountId y fechas
          let filterConditions = ["Status eq 'Actual'"];
          
          // Agregar filtro por cuenta si hay una seleccionada
          if (selectedAccountId) {
            filterConditions.push(`Account/Id eq ${selectedAccountId}`);
          }
          
          // Agregar filtro por fecha inicio si hay una seleccionada
          if (selectedDateFrom) {
            filterConditions.push(`ValueDate ge ${selectedDateFrom}`);
          }
          
          // Agregar filtro por fecha fin si hay una seleccionada
          if (selectedDateTo) {
            filterConditions.push(`ValueDate le ${selectedDateTo}`);
          }
          
          // Unir las condiciones del filtro con 'and'
          const filterParam = `$filter=${filterConditions.join(' and ')}`;
          params.push(filterParam);
          
          // Unir todos los parámetros
          // En desarrollo (webpack): usar ? y & normales
          // En producción (Vercel): codificar como %3F y %26
          if (params.length > 0) {
            const questionMark = isDevelopment ? '?' : '%3F';
            const ampersand = isDevelopment ? '&' : '%26';
            endpoint += questionMark + params.join(ampersand);
          }
          break;
      }
      
      // Agregar límite de registros para Cuentas y Flujos (codigos-presupuestarios, divisas y cotizacion ya lo gestionan dentro del switch)
      if (downloadType !== 'movimientos' && downloadType !== 'codigos-presupuestarios' && downloadType !== 'divisas' && downloadType !== 'cotizacion' && recordLimit !== 'all') {
        const hasParams = isDevelopment ? endpoint.includes('?') : endpoint.includes('%3F');
        const separator = hasParams 
          ? (isDevelopment ? '&' : '%26')
          : (isDevelopment ? '?' : '%3F');
        endpoint += separator + `$top=${recordLimit}`;
      }
      
      console.log("Descargando desde:", endpoint);
      console.log("Usuario autenticado:", userCredentials.username);
      
      // Intentar obtener los datos con autenticación
      let response;
      let retries = 3;
      while (retries > 0) {
        try {
          response = await authenticatedFetch(endpoint);
          console.log("Respuesta recibida. Status:", response.status);
          if (response.ok) break;
        } catch (fetchError) {
          console.error("Error en intento de fetch:", fetchError);
          retries--;
          if (retries === 0) throw new Error('Error al obtener datos después de 3 intentos');
          await new Promise(resolve => setTimeout(resolve, 1000)); // Esperar 1s antes de reintentar
        }
      }

      const data = await response.json();
      console.log("Datos recibidos:", data);


      // Verificar que tengamos datos
      if (!data || !data.value || data.value.length === 0) {
        showNotification("No se encontraron registros con los filtros seleccionados", "error");
        return; // Salir sin lanzar error
      }

      const records = data.value; // OData devuelve los datos en data.value
      
      // Determinar el nombre de la hoja según el tipo de descarga
      let sheetName = '';
      switch(downloadType) {
        case 'cuentas':
          sheetName = 'Accounts';
          break;
        case 'flujos':
          sheetName = 'Flujos';
          break;
        case 'codigos-presupuestarios':
          sheetName = 'Codigos Presupuestarios';
          break;
        case 'divisas':
          sheetName = 'Divisas';
          break;
        case 'cotizacion':
          sheetName = 'Cotizacion';
          break;
        case 'movimientos':
          sheetName = 'Movimientos';
          break;
        default:
          sheetName = downloadType;
      }
      
      // Verificar si la hoja existe y eliminarla
      try {
        const existingSheet = context.workbook.worksheets.getItem(sheetName);
        existingSheet.delete();
        await context.sync();
        console.log(`Hoja existente '${sheetName}' eliminada`);
      } catch (error) {
        // La hoja no existe, no hay problema
        console.log(`La hoja '${sheetName}' no existe, se creará una nueva`);
      }
      
      // Crear la hoja
      const sheet = context.workbook.worksheets.add(sheetName);
      sheet.load(["protection", "name"]);
      await context.sync();

      if (sheet.protection.protected) {
        throw new Error("La hoja está protegida. No se pueden escribir datos.");
      }
      
      console.log(`Hoja creada: ${sheetName}`);

      // Eliminar Sheet1 si existe (solo la primera vez)
      try {
        const sheet1 = context.workbook.worksheets.getItem("Sheet1");
        sheet1.delete();
        await context.sync();
        console.log("Hoja Sheet1 eliminada");
      } catch (error) {
        // Sheet1 no existe o ya fue eliminada, continuar normalmente
        console.log("Sheet1 no existe o ya fue eliminada");
      }

      /**
       * Formatea una fecha ISO a número serial de Excel
       * 
       * Excel almacena fechas como números seriales (días desde 30/12/1899)
       * Ejemplo: "2025-11-12" → 45962
       * 
       * @param {string} dateString - Fecha en formato ISO (YYYY-MM-DDTHH:mm:ss)
       * @param {string} fieldName - Nombre del campo (para logging)
       * @returns {number|string} Número serial de Excel o string vacío si la fecha es inválida
       */
      const formatDate = (dateString, fieldName) => {
        // Verificar si el valor es nulo, undefined o string vacío
        if (!dateString || dateString === '' || dateString === null || dateString === undefined) {
          console.log(`Campo ${fieldName}: valor vacío`);
          return '';
        }
        
        console.log(`Formateando ${fieldName}:`, dateString, 'Tipo:', typeof dateString);
        
        try {
          const date = new Date(dateString);
          
          // Verificar si la fecha es válida
          if (isNaN(date.getTime())) {
            console.warn(`Fecha inválida en ${fieldName}:`, dateString);
            return '';
          }
          
          // Convertir a número de serie de Excel
          // Excel cuenta los días desde 1/1/1900 (pero tiene un bug del año 1900)
          // JavaScript Date empieza desde 1/1/1970
          // Fórmula: (fecha en ms - fecha base) / ms por día + offset de Excel
          const excelEpoch = new Date(1899, 11, 30); // 30 de diciembre de 1899
          const msPerDay = 24 * 60 * 60 * 1000;
          const excelSerialDate = (date.getTime() - excelEpoch.getTime()) / msPerDay;
          
          console.log(`${fieldName} - Excel serial:`, excelSerialDate);
          return excelSerialDate;
        } catch (e) {
          console.error(`Error al formatear fecha ${fieldName}:`, dateString, e);
          return '';
        }
      };

      /**
       * Formatea un valor según el tipo de campo
       * 
       * Transformaciones especiales:
       * - @odata.etag: Se oculta (null)
       * - Booleanos: true/false → "true"/"false" (strings)
       * - Fechas: Se convierten a serial Excel
       * - ID: Se prefija con apóstrofe para forzar formato texto
       * - Otros: Se devuelven tal cual (null → '')
       * 
       * @param {string} fieldName - Nombre del campo
       * @param {any} value - Valor a formatear
       * @returns {any} Valor formateado según el tipo
       */
      const formatValue = (fieldName, value) => {
        // No mostrar @odata.etag
        if (fieldName === '@odata.etag') return null;

        // Formatear booleanos
        if (fieldName === 'Active' || fieldName === 'HasWarnings' || fieldName === 'IsInterco') {
          return value === true ? 'true' : value === false ? 'false' : '';
        }

        // Formatear fechas
        if (fieldName === 'CreationDateTime' || fieldName === 'ModificationDateTime' || 
            fieldName === 'BankClosingDate' || fieldName === 'CloseDate' || 
            fieldName === 'ValueDate' || fieldName === 'TrnDate') {
          return formatDate(value, fieldName);
        }

        // Convertir Id a String explícitamente con apóstrofe para forzar formato texto
        if (fieldName === 'Id') {
          // Agregar un espacio de ancho cero al inicio para forzar que Excel lo trate como texto
          return value !== undefined && value !== null ? "'" + String(value) : '';
        }

        // Para el resto de campos, devolver tal cual
        return value !== undefined && value !== null ? value : '';
      };

      // Preparar encabezados y datos según el tipo de descarga
      let headers = [];
      let values = [];

      if (downloadType === 'movimientos' && selectedFields.length > 0) {
        // Usar solo los campos seleccionados
        headers = selectedFields;
        values = records.map(record => 
          selectedFields.map(field => formatValue(field, record[field]))
        );
      } else {
        // Obtener todos los campos del primer registro, excluyendo @odata.etag
        const allFields = Object.keys(records[0]).filter(key => key !== '@odata.etag');
        headers = allFields;
        
        values = records.map(record => 
          allFields.map(field => formatValue(field, record[field]))
        );
      }

      // Calcular el rango necesario
      const numRows = values.length + 1; // +1 para la fila de encabezados
      const numCols = headers.length;
      
      /**
       * Convierte un índice numérico a letra de columna Excel
       * 
       * Sistema de base 26:
       * 0 → A, 1 → B, ... 25 → Z
       * 26 → AA, 27 → AB, ... 51 → AZ
       * 52 → BA, etc.
       * 
       * @param {number} colIndex - Índice de columna (base 0)
       * @returns {string} Letra de columna Excel
       */
      const getColumnLetter = (colIndex) => {
        let letter = '';
        while (colIndex >= 0) {
          letter = String.fromCharCode((colIndex % 26) + 65) + letter;
          colIndex = Math.floor(colIndex / 26) - 1;
        }
        return letter;
      };
      const endColumn = getColumnLetter(numCols - 1);
      
      // Escribir datos en un solo bloque para mejor rendimiento
      const range = sheet.getRange(`A1:${endColumn}${numRows}`);
      range.values = [headers, ...values];

      // Aplicar formato en una sola operación
      const headerRange = range.getRow(0);
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.bold = true;
      headerRange.format.font.color = "#FFFFFF";

      // Aplicar formato de fecha a las columnas de fecha
      const dateFields = ['CreationDateTime', 'ModificationDateTime', 'BankClosingDate', 'CloseDate', 'ValueDate', 'TrnDate'];
      dateFields.forEach(dateField => {
        const colIndex = headers.indexOf(dateField);
        if (colIndex >= 0) {
          const colLetter = getColumnLetter(colIndex);
          const dateRange = sheet.getRange(`${colLetter}2:${colLetter}${numRows}`);
          dateRange.numberFormat = [["DD/MM/YYYY"]];
          console.log(`Formato de fecha aplicado a columna ${colLetter} (${dateField})`);
        }
      });

      // Aplicar formato de texto a la columna Id para evitar notación científica
      const idColIndex = headers.indexOf('Id');
      if (idColIndex >= 0) {
        const idColLetter = getColumnLetter(idColIndex);
        const idRange = sheet.getRange(`${idColLetter}2:${idColLetter}${numRows}`);
        idRange.numberFormat = [["@"]]; // @ significa formato texto en Excel
        console.log(`Formato de texto aplicado a columna ${idColLetter} (Id)`);
      }

      // Autoajustar columnas
      range.format.autofitColumns();
      
      // Activar la hoja para que el foco se quede en ella
      sheet.activate();
      
      await context.sync();
      showNotification(`¡${records.length} ${downloadType} descargados exitosamente!`, "success");
    });
  } catch (error) {
    console.error("Error específico:", error.message);
    let errorMessage = "Error al descargar los datos";
    
    // Mensajes de error más específicos
    if (error.message.includes("protegida")) {
      errorMessage = "La hoja está protegida. Desproteja la hoja e intente nuevamente.";
    } else if (error.message.includes("obtener datos")) {
      errorMessage = "Error de conexión. Verifique su conexión a internet.";
    }
    
    showNotification(errorMessage, "error");
  }
}

/**
 * Muestra el modal de configuración de importación
 * 
 * Permite al usuario seleccionar:
 * - Tipo de importación: Flujos o Movimientos
 * 
 * @async
 * @throws {Error} Si el usuario no está autenticado o hay problemas con el DOM
 */
export async function importData() {
  try {
    // Verificar que el usuario esté logueado
    if (!userCredentials.isLoggedIn) {
      showNotification("Debe iniciar sesión primero", "error");
      return;
    }

    const modal = document.getElementById("importModal");
    modal.classList.remove("hidden");
    modal.style.display = "block";

    // Configurar botón de crear hoja
    document.getElementById("importCreateSheet").onclick = async () => {
      await executeCreateSheet();
    };

    // Configurar botón de submit
    document.getElementById("importSubmit").onclick = async () => {
      await executeImport();
    };

    // Configurar botón de cancelar
    document.getElementById("importCancel").onclick = () => {
      modal.classList.add("hidden");
    };

    // Cerrar modal al hacer clic fuera
    window.onclick = (event) => {
      if (event.target === modal) {
        modal.classList.add("hidden");
      }
    };
  } catch (error) {
    console.error("Error al abrir modal de importación:", error);
    showNotification("Error al abrir el modal de importación", "error");
  }
}

/**
 * Crea una hoja de Excel con las cabeceras según el tipo seleccionado
 * 
 * @async
 * @throws {Error} Si no se seleccionó un tipo o hay problemas al crear la hoja
 */
async function executeCreateSheet() {
  try {
    const importType = document.getElementById("importType").value;
    const importError = document.getElementById("importError");
    
    // Validar que se haya seleccionado una opción
    if (!importType) {
      importError.textContent = "Debe seleccionar un tipo de importación";
      importError.classList.remove("hidden");
      return;
    }

    // Ocultar mensaje de error si había alguno
    importError.classList.add("hidden");

    console.log("Creando hoja para tipo:", importType);

    // Cerrar el modal
    document.getElementById("importModal").classList.add("hidden");

    // Crear la hoja según el tipo
    if (importType === "movimientos") {
      // Cargar todos los datos necesarios antes de crear la hoja
      showNotification("Descargando datos necesarios...", "info");
      await Promise.all([
        loadAccounts(),
        loadFlowCodes(),
        loadBudgetCodes(),
        loadCurrencies(),
        loadQuotationPlaces()
      ]);
      await createMovimientosSheet();
    } else if (importType === "flujos") {
      // TODO: Implementar creación de hoja para flujos
      showNotification(`Funcionalidad de creación de hoja para flujos en desarrollo`, "info");
    }
  } catch (error) {
    console.error("Error en executeCreateSheet:", error);
    showNotification("Error al crear la hoja", "error");
  }
}

/**
 * Recopila las opciones seleccionadas del modal y ejecuta la importación
 * 
 * Validaciones:
 * - Verifica que se haya seleccionado un tipo de importación
 * 
 * @async
 * @throws {Error} Si no se cumplen las validaciones o hay problemas al preparar la importación
 */
async function executeImport() {
  try {
    const importType = document.getElementById("importType").value;
    const importError = document.getElementById("importError");
    
    // Validar que se haya seleccionado una opción
    if (!importType) {
      importError.textContent = "Debe seleccionar un tipo de importación";
      importError.classList.remove("hidden");
      return;
    }

    // Ocultar mensaje de error si había alguno
    importError.classList.add("hidden");

    console.log("Importación de tipo:", importType);

    // Cerrar el modal
    document.getElementById("importModal").classList.add("hidden");

    // Ejecutar importación según el tipo
    if (importType === "movimientos") {
      await importMovimientosToOData();
    } else if (importType === "flujos") {
      showNotification(`Funcionalidad de importación de flujos en desarrollo`, "info");
    }
  } catch (error) {
    console.error("Error en executeImport:", error);
    showNotification("Error al preparar la importación", "error");
  }
}

/**
 * Lee los datos de la hoja "Movimientos" en Excel
 * 
 * @async
 * @returns {Promise<Array<Object>>} Array de objetos con los datos de cada fila
 * @throws {Error} Si la hoja no existe o hay problemas al leer los datos
 */
async function readMovimientosSheet() {
  return await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getItem("Movimientos");
      const usedRange = sheet.getUsedRange();
      usedRange.load(["values", "rowCount"]);
      await context.sync();

      const values = usedRange.values;
      
      if (values.length <= 1) {
        throw new Error("La hoja no contiene datos para importar");
      }

      // La primera fila son las cabeceras
      const headers = values[0];
      const records = [];

      // Procesar cada fila de datos (empezando desde la fila 2)
      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        const record = {};
        
        // Mapear cada columna a su campo correspondiente
        for (let j = 0; j < headers.length; j++) {
          const header = headers[j];
          const value = row[j];
          
          // Saltar valores vacíos para campos opcionales
          if (value !== null && value !== undefined && value !== "") {
            record[header] = value;
          }
        }
        
        // Solo agregar registros que tengan al menos un campo
        if (Object.keys(record).length > 0) {
          records.push(record);
        }
      }

      console.log(`Leídos ${records.length} registros de la hoja Movimientos`);
      return records;
    } catch (error) {
      if (error.message.includes("ItemNotFound")) {
        throw new Error("No existe la hoja 'Movimientos'. Debe crearla primero.");
      }
      throw error;
    }
  });
}

/**
 * Valida un registro de movimiento según los criterios de la Historia de Usuario
 * 
 * @param {Object} record - Registro a validar
 * @param {number} rowNumber - Número de fila (para mensajes de error)
 * @returns {Object} Objeto con {isValid: boolean, errors: string[], errorFields: string[]}
 */
function validateMovimientoRecord(record, rowNumber) {
  const errors = [];
  const errorFields = []; // Campos con error para marcar en rojo

  // Validar Status (requerido)
  if (!record.Status || record.Status.toString().trim() === "") {
    errors.push(`Fila ${rowNumber}: El campo Status es obligatorio`);
    errorFields.push('Status');
  }

  // Validar IsDebit (requerido)
  if (record.IsDebit === null || record.IsDebit === undefined || record.IsDebit === "") {
    errors.push(`Fila ${rowNumber}: El campo IsDebit es obligatorio`);
    errorFields.push('IsDebit');
  }

  // Validar Amount (distinto de 0)
  if (!record.Amount || parseFloat(record.Amount) === 0) {
    errors.push(`Fila ${rowNumber}: El campo Amount debe ser distinto de 0`);
    errorFields.push('Amount');
  }

  // Validar ValueDate (requerido y formato válido)
  if (!record.ValueDate) {
    errors.push(`Fila ${rowNumber}: El campo ValueDate es obligatorio`);
    errorFields.push('ValueDate');
  } else if (!isValidDate(record.ValueDate)) {
    errors.push(`Fila ${rowNumber}: El campo ValueDate tiene formato inválido. Use dd/mm/yyyy`);
    errorFields.push('ValueDate');
  }

  // Validar TrnAmount (distinto de 0)
  if (!record.TrnAmount || parseFloat(record.TrnAmount) === 0) {
    errors.push(`Fila ${rowNumber}: El campo TrnAmount debe ser distinto de 0`);
    errorFields.push('TrnAmount');
  }

  // Validar TrnDate (requerido y formato válido)
  if (!record.TrnDate) {
    errors.push(`Fila ${rowNumber}: El campo TrnDate es obligatorio`);
    errorFields.push('TrnDate');
  } else if (!isValidDate(record.TrnDate)) {
    errors.push(`Fila ${rowNumber}: El campo TrnDate tiene formato inválido. Use dd/mm/yyyy`);
    errorFields.push('TrnDate');
  }

  // Validar Number (mayor o igual a 1)
  if (!record.Number || parseInt(record.Number) < 1) {
    errors.push(`Fila ${rowNumber}: El campo Number debe ser >= 1`);
    errorFields.push('Number');
  }

  // Validar Account (requerido)
  if (!record.Account || record.Account.toString().trim() === "") {
    errors.push(`Fila ${rowNumber}: El campo Account es obligatorio`);
    errorFields.push('Account');
  }

  // Validar campos booleanos requeridos
  const booleanFields = ['UseInBalanceVal', 'UseInBalanceTrn', 'Interco', 'UseIntercoChart', 'IsManualFee'];
  booleanFields.forEach(field => {
    if (record[field] === null || record[field] === undefined || record[field] === "") {
      errors.push(`Fila ${rowNumber}: El campo ${field} es obligatorio`);
      errorFields.push(field);
    }
  });

  return {
    isValid: errors.length === 0,
    errors: errors,
    errorFields: errorFields,
    rowNumber: rowNumber
  };
}

/**
 * Marca las celdas con errores de validación en rojo
 * 
 * @async
 * @param {Array<Object>} validationResults - Resultados de validación con errorFields y rowNumber
 */
async function markErrorCells(validationResults) {
  await Excel.run(async (context) => {
    try {
      const sheet = context.workbook.worksheets.getItem("Movimientos");
      
      // Obtener las cabeceras para saber qué columna es cada campo
      const headerRange = sheet.getRange("A1:S1");
      headerRange.load("values");
      await context.sync();
      
      const headers = headerRange.values[0];
      
      // Procesar cada resultado de validación que tenga errores
      validationResults.forEach(result => {
        if (!result.isValid && result.errorFields.length > 0) {
          result.errorFields.forEach(fieldName => {
            // Encontrar el índice de la columna para este campo
            const columnIndex = headers.indexOf(fieldName);
            
            if (columnIndex !== -1) {
              // Convertir índice de columna a letra (A, B, C, etc.)
              const columnLetter = String.fromCharCode(65 + columnIndex);
              const cellAddress = `${columnLetter}${result.rowNumber}`;
              
              // Marcar solo el contorno de la celda en rojo
              const errorCell = sheet.getRange(cellAddress);
              ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(edge => {
                const border = errorCell.format.borders.getItem(edge);
                border.style = "Continuous";
                border.color = "#CC0000"; // Borde rojo
              });
            }
          });
        }
      });
      
      await context.sync();
      console.log("Celdas con errores marcadas con borde rojo");
    } catch (error) {
      console.error("Error al marcar celdas:", error);
    }
  });
}

/**
 * Valida si un valor es una fecha válida en formato dd/mm/yyyy o número serial de Excel
 * 
 * @param {any} value - Valor a validar
 * @returns {boolean} true si es una fecha válida
 */
function isValidDate(value) {
  if (!value) return false;

  // Si es un número (serial de Excel), verificar que esté en rango válido
  if (typeof value === 'number') {
    return value > 0 && value < 2958466; // Rango válido de Excel (1900-9999)
  }

  // Si es string, validar formato dd/mm/yyyy
  if (typeof value === 'string') {
    const datePattern = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    const match = value.match(datePattern);
    
    if (!match) return false;
    
    const day = parseInt(match[1]);
    const month = parseInt(match[2]);
    const year = parseInt(match[3]);
    
    // Validar rangos
    if (month < 1 || month > 12) return false;
    if (day < 1 || day > 31) return false;
    if (year < 1900 || year > 9999) return false;
    
    // Validar días por mes
    const daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    const isLeapYear = (year % 4 === 0 && year % 100 !== 0) || (year % 400 === 0);
    if (month === 2 && isLeapYear) {
      return day <= 29;
    }
    return day <= daysInMonth[month - 1];
  }

  return false;
}

/**
 * Convierte una fecha de Excel (número serial o dd/mm/yyyy) a formato ISO 8601
 * 
 * @param {any} dateValue - Valor de fecha (número serial de Excel o string dd/mm/yyyy)
 * @returns {string} Fecha en formato ISO 8601 (YYYY-MM-DDTHH:mm:ssZ)
 */
function convertToISO8601(dateValue) {
  let date;

  // Si es un número serial de Excel
  if (typeof dateValue === 'number') {
    // Excel cuenta los días desde 30/12/1899
    const excelEpoch = new Date(1899, 11, 30);
    const msPerDay = 24 * 60 * 60 * 1000;
    date = new Date(excelEpoch.getTime() + dateValue * msPerDay);
  } 
  // Si es string en formato dd/mm/yyyy
  else if (typeof dateValue === 'string') {
    const parts = dateValue.split('/');
    const day = parseInt(parts[0]);
    const month = parseInt(parts[1]) - 1; // JavaScript months son 0-indexed
    const year = parseInt(parts[2]);
    date = new Date(year, month, day);
  }
  else {
    throw new Error(`Formato de fecha no soportado: ${dateValue}`);
  }

  // Convertir a ISO 8601 (formato UTC con Z)
  return date.toISOString().split('.')[0] + 'Z';
}

/**
 * Construye el payload JSON para enviar un movimiento al servidor OData
 * 
 * @param {Object} record - Registro con los datos del movimiento
 * @returns {Object} Objeto JSON en el formato requerido por OData
 */
function buildMovimientoJSON(record) {
  const payload = {};

  // Campo opcional TERCERO - solo se incluye si tiene valor
  if (record.TERCERO && record.TERCERO.toString().trim() !== "") {
    payload.TERCERO = record.TERCERO.toString().trim();
  }

  // Construir el objeto Entity
  payload.Entity = {
    Status: record.Status.toString(),
    IsDebit: parseBooleanField(record.IsDebit),
    Amount: parseFloat(record.Amount),
    ValueDate: convertToISO8601(record.ValueDate),
    TrnAmount: parseFloat(record.TrnAmount),
    TrnDate: convertToISO8601(record.TrnDate),
    Number: parseInt(record.Number),
    UseInBalanceVal: parseBooleanField(record.UseInBalanceVal),
    UseInBalanceTrn: parseBooleanField(record.UseInBalanceTrn),
    Interco: parseBooleanField(record.Interco),
    UseIntercoChart: parseBooleanField(record.UseIntercoChart),
    IsManualFee: parseBooleanField(record.IsManualFee)
  };

  // Descripción (opcional)
  if (record.Description) {
    payload.Entity.Description = record.Description.toString();
  }

  // Mapear Account a ID usando el caché
  const accountCode = record.Account.toString().trim();
  const account = cachedAccounts.find(acc => acc.Code === accountCode);
  if (account) {
    payload.Entity["Account@odata.bind"] = `Account2CashSet(${account.Id})`;
  } else {
    throw new Error(`No se encontró la cuenta con código: ${accountCode}`);
  }

  // Mapear BudgetCode a ID (opcional)
  if (record.BudgetCode) {
    const budgetCode = record.BudgetCode.toString().trim();
    const budget = cachedBudgetCodes.find(bc => bc.Code === budgetCode);
    if (budget) {
      payload.Entity["BudgetCode@odata.bind"] = `BudgetCodeSet(${budget.Id})`;
    } else {
      console.warn(`No se encontró el código presupuestario: ${budgetCode}`);
    }
  }

  // Mapear FlowCode a ID (opcional)
  if (record.FlowCode) {
    const flowCode = record.FlowCode.toString().trim();
    const flow = cachedFlowCodes.find(fc => fc.Code === flowCode);
    if (flow) {
      payload.Entity["FlowCode@odata.bind"] = `FlowCodeSet(${flow.Id})`;
    } else {
      console.warn(`No se encontró el flujo de caja: ${flowCode}`);
    }
  }

  // Mapear TrnCurrency a ID (opcional)
  if (record.TrnCurrency) {
    const currencyCode = record.TrnCurrency.toString().trim();
    const currency = cachedCurrencies.find(c => c.Code === currencyCode);
    if (currency) {
      payload.Entity["TrnCurrency@odata.bind"] = `CurrencySet('${currency.Id}')`;
    } else {
      console.warn(`No se encontró la divisa: ${currencyCode}`);
    }
  }

  // Mapear QuotationPlace a ID (opcional)
  if (record.QuotationPlace) {
    const quotationDesc = record.QuotationPlace.toString().trim();
    const quotation = cachedQuotationPlaces.find(qp => qp.Description === quotationDesc);
    if (quotation) {
      payload.Entity["QuotationPlace@odata.bind"] = `QuotationPlaceSet(${quotation.Id})`;
    } else {
      console.warn(`No se encontró el lugar de cotización: ${quotationDesc}`);
    }
  }

  return payload;
}

/**
 * Parsea un campo booleano de Excel a booleano JavaScript
 * 
 * @param {any} value - Valor del campo (puede ser string "true"/"false" o booleano)
 * @returns {boolean} Valor booleano
 */
function parseBooleanField(value) {
  if (typeof value === 'boolean') return value;
  if (typeof value === 'string') {
    return value.toLowerCase() === 'true';
  }
  return Boolean(value);
}

/**
 * Construye un JSON de batch request para OData v4
 * 
 * @param {Array} records - Array de registros de movimientos
 * @returns {Object} Objeto con la estructura de batch request
 */
function buildBatchRequestJSON(records) {
  const requests = records.map((record, index) => {
    // Construir el payload individual usando la función existente
    const payload = buildMovimientoJSON(record);
    
    return {
      id: String(index + 1),
      method: "POST",
      url: "CashFlowDtoWithExtendersSet",
      headers: {
        "Content-Type": "application/json"
      },
      body: payload
    };
  });
  
  return { requests };
}

/**
 * Importa movimientos desde la hoja de Excel al servidor OData
 * 
 * @async
 */
async function importMovimientosToOData() {
  try {
    showNotification("Preparando importación...", "info");

    // 1. Cargar todos los datos de referencia necesarios
    console.log("Cargando datos de referencia...");
    await Promise.all([
      loadAccounts(),
      loadFlowCodes(),
      loadBudgetCodes(),
      loadCurrencies(),
      loadQuotationPlaces()
    ]);

    // 2. Leer datos de la hoja
    console.log("Leyendo datos de la hoja Movimientos...");
    const records = await readMovimientosSheet();

    if (records.length === 0) {
      showNotification("No hay datos para importar", "error");
      return;
    }

    // 3. Validar todos los registros
    console.log(`Validando ${records.length} registros...`);
    const validationResults = records.map((record, index) => 
      validateMovimientoRecord(record, index + 2) // +2 porque la fila 1 son cabeceras
    );

    const allErrors = validationResults.flatMap(result => result.errors);
    
    if (allErrors.length > 0) {
      console.error("Errores de validación:", allErrors);
      
      // Marcar celdas con errores en rojo
      await markErrorCells(validationResults);
      
      // Crear mensaje detallado con los campos problemáticos
      const invalidRecords = validationResults.filter(r => !r.isValid);
      let errorMessage = `⚠️ Validación fallida: ${allErrors.length} error(es) encontrado(s)\n\n`;
      
      invalidRecords.forEach(result => {
        errorMessage += `📍 Fila ${result.rowNumber}:\n`;
        errorMessage += `   Campos con problema: ${result.errorFields.join(', ')}\n\n`;
      });
      
      errorMessage += "Las celdas con errores han sido marcadas en rojo. Corríjalas e intente de nuevo.";
      
      // Mostrar mensaje en notificación
      showNotification("Errores de validación encontrados", "error");
      
      // Mostrar mensaje detallado en panel modal
      showErrorDetails(errorMessage);
      
      allErrors.forEach(error => console.error(error));
      return;
    }

    // 4. Si solo hay un registro, enviarlo
    if (records.length === 1) {
      console.log("Enviando único registro al servidor...");
      const payload = buildMovimientoJSON(records[0]);

      console.log("Payload JSON:", JSON.stringify(payload, null, 2));

      const isDevelopment = window.location.hostname === 'localhost';
      const VERCEL_PROXY = isDevelopment
        ? '/odata/'
        : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';

      const endpoint = `${VERCEL_PROXY}CashFlowDtoWithExtendersSet`;

      console.log("Enviando POST a:", endpoint);

      const response = await authenticatedFetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Accept': 'application/json; charset=utf-8'
        },
        body: JSON.stringify(payload)
      });

      if (response.ok) {
        const result = await response.json();
        console.log("Respuesta del servidor:", result);
        showNotification("✅ Movimiento importado exitosamente al servidor OData", "success");
      } else {
        const errorText = await response.text();
        console.error("Error del servidor:", response.status, errorText);
        
        // Crear mensaje de error detallado
        let errorMessage = `❌ Error al añadir el movimiento\n\n`;
        errorMessage += `Código de error: ${response.status}\n`;
        errorMessage += `Detalles: ${errorText.substring(0, 200)}\n\n`;
        errorMessage += `💡 Sugerencias:\n`;
        errorMessage += `- Verifique que el campo Status sea "Actual" (no "Active")\n`;
        errorMessage += `- Revise que todos los códigos de cuenta, flujo y presupuesto sean válidos\n`;
        errorMessage += `- Compruebe que las fechas estén en formato correcto\n`;
        
        showNotification("Error al importar movimiento", "error");
        showErrorDetails(errorMessage);
      }
    } else {
      // Múltiples registros: usar OData $batch
      console.log(`Enviando ${records.length} registros en lote al servidor...`);
      const batchPayload = buildBatchRequestJSON(records);

      console.log("Batch Payload JSON:", JSON.stringify(batchPayload, null, 2));

      const isDevelopment = window.location.hostname === 'localhost';
      const VERCEL_PROXY = isDevelopment
        ? '/odata/'
        : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';

      const endpoint = `${VERCEL_PROXY}$batch`;

      console.log("Enviando POST batch a:", endpoint);

      const response = await authenticatedFetch(endpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json; charset=utf-8',
          'Accept': 'application/json; charset=utf-8'
        },
        body: JSON.stringify(batchPayload)
      });

      if (response.ok) {
        const result = await response.json();
        console.log("Respuesta del servidor (batch):", result);
        
        // Analizar respuesta batch para ver cuántos tuvieron éxito
        const responses = result.responses || [];
        const successCount = responses.filter(r => r.status >= 200 && r.status < 300).length;
        const errorCount = responses.length - successCount;
        
        if (errorCount === 0) {
          showNotification(`✅ ${successCount} movimientos importados exitosamente`, "success");
        } else {
          let errorMessage = `⚠️ Importación parcial:\n\n`;
          errorMessage += `✅ Exitosos: ${successCount}\n`;
          errorMessage += `❌ Fallidos: ${errorCount}\n\n`;
          errorMessage += `Detalles de errores:\n`;
          
          responses.forEach((resp, idx) => {
            if (resp.status >= 300) {
              errorMessage += `\n• Registro ${idx + 1} (fila ${idx + 2}): Error ${resp.status}\n`;
              if (resp.body && resp.body.error) {
                errorMessage += `  ${resp.body.error.message}\n`;
              }
            }
          });
          
          showNotification(`Importación completada con errores`, "warning");
          showErrorDetails(errorMessage);
        }
      } else {
        const errorText = await response.text();
        console.error("Error del servidor (batch):", response.status, errorText);
        
        let errorMessage = `❌ Error al enviar el lote de movimientos\n\n`;
        errorMessage += `Código de error: ${response.status}\n`;
        errorMessage += `Detalles: ${errorText.substring(0, 300)}\n\n`;
        errorMessage += `💡 Sugerencias:\n`;
        errorMessage += `- Verifique que todos los registros tengan datos válidos\n`;
        errorMessage += `- Compruebe la conectividad con el servidor OData\n`;
        errorMessage += `- Revise los logs de consola para más detalles\n`;
        
        showNotification("Error al importar lote de movimientos", "error");
        showErrorDetails(errorMessage);
      }
    }

  } catch (error) {
    console.error("Error durante la importación:", error);
    showNotification("Error durante la importación: " + error.message, "error");
  }
}

/**
 * Crea una hoja de Excel para importar Movimientos con las cabeceras predefinidas
 * 
 * @async
 * @throws {Error} Si hay problemas al crear la hoja o escribir las cabeceras
 */
async function createMovimientosSheet() {
  try {
    await Excel.run(async (context) => {
      const application = context.workbook.application;
      application.suspendScreenUpdatingUntilNextSync();
      
      const sheetName = "Movimientos";
      
      // Verificar si la hoja existe y eliminarla
      try {
        const existingSheet = context.workbook.worksheets.getItem(sheetName);
        existingSheet.delete();
        await context.sync();
        console.log(`Hoja existente '${sheetName}' eliminada`);
      } catch (error) {
        console.log(`La hoja '${sheetName}' no existe, se creará una nueva`);
      }
      
      // Crear la hoja
      const sheet = context.workbook.worksheets.add(sheetName);
      sheet.load(["protection", "name"]);
      await context.sync();

      if (sheet.protection.protected) {
        throw new Error("La hoja está protegida. No se pueden escribir datos.");
      }
      
      console.log(`Hoja creada: ${sheetName}`);

      // Definir las cabeceras basadas en el JSON
      const headers = [
        "TERCERO",
        "Status",
        "IsDebit",
        "Amount",
        "ValueDate",
        "TrnAmount",
        "TrnDate",
        "Number",
        "Description",
        "Account",
        "BudgetCode",
        "FlowCode",
        "TrnCurrency",
        "QuotationPlace",
        "UseInBalanceVal",
        "UseInBalanceTrn",
        "Interco",
        "UseIntercoChart",
        "IsManualFee"
      ];

      // Escribir las cabeceras en la primera fila
      const headerRange = sheet.getRange(`A1:${String.fromCharCode(64 + headers.length)}1`);
      headerRange.values = [headers];
      
      // Formatear las cabeceras
      headerRange.format.fill.color = "#4472C4";
      headerRange.format.font.color = "white";
      headerRange.format.font.bold = true;
      headerRange.format.horizontalAlignment = "Center";
      
      // Autoajustar columnas
      sheet.getUsedRange().format.autofitColumns();
      
      // Agregar validación de datos para las columnas de fecha
      // ValueDate está en la columna E (índice 4)
      const valueDateColumn = sheet.getRange("E2:E1048576");
      const valueDateValidation = valueDateColumn.dataValidation;
      valueDateValidation.rule = {
        date: {
          operator: Excel.DataValidationOperator.greaterThan,
          formula1: "1900-01-01"
        }
      };
      valueDateValidation.prompt = {
        message: "Ingrese una fecha válida (formato: dd/mm/yyyy)",
        showPrompt: true,
        title: "Fecha de valor"
      };
      valueDateValidation.errorAlert = {
        message: "Debe ingresar una fecha válida",
        showAlert: true,
        style: Excel.DataValidationAlertStyle.stop,
        title: "Fecha inválida"
      };
      
      // TrnDate está en la columna G (índice 6)
      const trnDateColumn = sheet.getRange("G2:G1048576");
      const trnDateValidation = trnDateColumn.dataValidation;
      trnDateValidation.rule = {
        date: {
          operator: Excel.DataValidationOperator.greaterThan,
          formula1: "1900-01-01"
        }
      };
      trnDateValidation.prompt = {
        message: "Ingrese una fecha válida (formato: dd/mm/yyyy)",
        showPrompt: true,
        title: "Fecha de transacción"
      };
      trnDateValidation.errorAlert = {
        message: "Debe ingresar una fecha válida",
        showAlert: true,
        style: Excel.DataValidationAlertStyle.stop,
        title: "Fecha inválida"
      };
      
      // Agregar dropdown para la columna Account (columna J, índice 9)
      if (cachedAccounts && cachedAccounts.length > 0) {
        const accountColumn = sheet.getRange("J2:J1048576");
        const accountValidation = accountColumn.dataValidation;
        
        // Crear la lista de valores separados por comas (solo los códigos de cuenta)
        const accountCodes = cachedAccounts.map(acc => acc.Code).join(",");
        
        accountValidation.rule = {
          list: {
            inCellDropDown: true,
            source: accountCodes
          }
        };
        accountValidation.prompt = {
          message: "Seleccione una cuenta de la lista",
          showPrompt: true,
          title: "Cuenta contable"
        };
        accountValidation.errorAlert = {
          message: "Debe seleccionar una cuenta válida de la lista",
          showAlert: true,
          style: Excel.DataValidationAlertStyle.stop,
          title: "Cuenta inválida"
        };
        
        console.log(`Dropdown de cuentas configurado con ${cachedAccounts.length} cuentas`);
      } else {
        console.warn("No hay cuentas en caché. Descargue las cuentas primero para habilitar el dropdown.");
      }
      
      // Agregar dropdown para la columna BudgetCode (columna K, índice 10)
      if (cachedBudgetCodes && cachedBudgetCodes.length > 0) {
        const budgetCodeColumn = sheet.getRange("K2:K1048576");
        const budgetCodeValidation = budgetCodeColumn.dataValidation;
        
        const budgetCodes = cachedBudgetCodes.map(bc => bc.Code).join(",");
        
        budgetCodeValidation.rule = {
          list: {
            inCellDropDown: true,
            source: budgetCodes
          }
        };
        budgetCodeValidation.prompt = {
          message: "Seleccione un código presupuestario de la lista",
          showPrompt: true,
          title: "Código presupuestario"
        };
        budgetCodeValidation.errorAlert = {
          message: "Debe seleccionar un código presupuestario válido de la lista",
          showAlert: true,
          style: Excel.DataValidationAlertStyle.warning,
          title: "Código presupuestario inválido"
        };
        
        console.log(`Dropdown de códigos presupuestarios configurado con ${cachedBudgetCodes.length} códigos`);
      }
      
      // Agregar dropdown para la columna FlowCode (columna L, índice 11)
      if (cachedFlowCodes && cachedFlowCodes.length > 0) {
        const flowCodeColumn = sheet.getRange("L2:L1048576");
        const flowCodeValidation = flowCodeColumn.dataValidation;
        
        const flowCodes = cachedFlowCodes.map(fc => fc.Code).join(",");
        
        flowCodeValidation.rule = {
          list: {
            inCellDropDown: true,
            source: flowCodes
          }
        };
        flowCodeValidation.prompt = {
          message: "Seleccione un flujo de caja de la lista",
          showPrompt: true,
          title: "Flujo de caja"
        };
        flowCodeValidation.errorAlert = {
          message: "Debe seleccionar un flujo de caja válido de la lista",
          showAlert: true,
          style: Excel.DataValidationAlertStyle.warning,
          title: "Flujo de caja inválido"
        };
        
        console.log(`Dropdown de flujos de caja configurado con ${cachedFlowCodes.length} flujos`);
      }
      
      // Agregar dropdown para la columna TrnCurrency (columna M, índice 12)
      if (cachedCurrencies && cachedCurrencies.length > 0) {
        const currencyColumn = sheet.getRange("M2:M1048576");
        const currencyValidation = currencyColumn.dataValidation;
        
        // Crear la lista de códigos de divisa separados por comas
        const currencyCodes = cachedCurrencies.map(c => c.Code).join(",");
        
        currencyValidation.rule = {
          list: {
            inCellDropDown: true,
            source: currencyCodes
          }
        };
        currencyValidation.prompt = {
          message: "Seleccione una divisa de la lista",
          showPrompt: true,
          title: "Divisa de transacción"
        };
        currencyValidation.errorAlert = {
          message: "Debe seleccionar una divisa válida",
          showAlert: true,
          style: Excel.DataValidationAlertStyle.warning,
          title: "Divisa inválida"
        };
        
        console.log(`Dropdown de divisas configurado con ${cachedCurrencies.length} divisas`);
      } else {
        console.warn("No hay divisas en caché. Descargue las divisas primero para habilitar el dropdown.");
      }
      
      // Agregar dropdown para la columna QuotationPlace (columna N, índice 13)
      if (cachedQuotationPlaces && cachedQuotationPlaces.length > 0) {
        const quotationPlaceColumn = sheet.getRange("N2:N1048576");
        const quotationPlaceValidation = quotationPlaceColumn.dataValidation;
        
        // Crear la lista de descripciones de lugares de cotización separadas por comas
        const quotationDescriptions = cachedQuotationPlaces.map(qp => qp.Description).join(",");
        
        quotationPlaceValidation.rule = {
          list: {
            inCellDropDown: true,
            source: quotationDescriptions
          }
        };
        quotationPlaceValidation.prompt = {
          message: "Seleccione una opción de cotización",
          showPrompt: true,
          title: "Lugar de cotización"
        };
        quotationPlaceValidation.errorAlert = {
          message: "Debe seleccionar un lugar de cotización válido",
          showAlert: true,
          style: Excel.DataValidationAlertStyle.warning,
          title: "Lugar de cotización inválido"
        };
        
        console.log(`Dropdown de lugares de cotización configurado con ${cachedQuotationPlaces.length} lugares`);
      } else {
        console.warn("No hay lugares de cotización en caché. Descargue primero para habilitar el dropdown.");
      }
      
      // Agregar dropdown para Status (columna B, índice 1)
      const statusColumn = sheet.getRange("B2:B1048576");
      const statusValidation = statusColumn.dataValidation;
      statusValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "Actual,Forecast,Historical"
        }
      };
      statusValidation.prompt = {
        message: "Seleccione el estado del movimiento (Actual es el más común)",
        showPrompt: true,
        title: "Estado del movimiento"
      };
      statusValidation.errorAlert = {
        message: "Debe seleccionar un estado válido: Actual, Forecast o Historical",
        showAlert: true,
        style: Excel.DataValidationAlertStyle.stop,
        title: "Estado inválido"
      };
      console.log("Dropdown de Status configurado con valores: Actual, Forecast, Historical");

      // Agregar dropdowns booleanos (true/false) para múltiples columnas
      const booleanColumns = [
        { range: "C2:C1048576", name: "IsDebit", title: "¿Es débito?" },
        { range: "O2:O1048576", name: "UseInBalanceVal", title: "¿Usar en balance de valor?" },
        { range: "P2:P1048576", name: "UseInBalanceTrn", title: "¿Usar en balance de transacción?" },
        { range: "Q2:Q1048576", name: "Interco", title: "¿Es intercompañía?" },
        { range: "R2:R1048576", name: "UseIntercoChart", title: "¿Usar plan intercompañía?" },
        { range: "S2:S1048576", name: "IsManualFee", title: "¿Es comisión manual?" }
      ];
      
      booleanColumns.forEach(col => {
        const column = sheet.getRange(col.range);
        const validation = column.dataValidation;
        validation.rule = {
          list: {
            inCellDropDown: true,
            source: "true,false"
          }
        };
        validation.prompt = {
          message: "Seleccione true o false",
          showPrompt: true,
          title: col.title
        };
        validation.errorAlert = {
          message: "Debe seleccionar true o false",
          showAlert: true,
          style: Excel.DataValidationAlertStyle.stop,
          title: "Valor inválido"
        };
      });
      
      console.log("Todas las validaciones de datos configuradas correctamente");
      
      // Activar la hoja
      sheet.activate();
      
      await context.sync();
      
      showNotification("Hoja 'Movimientos' creada con controles de validación", "success");
    });
  } catch (error) {
    console.error("Error al crear hoja de Movimientos:", error);
    showNotification("Error al crear la hoja de Movimientos: " + error.message, "error");
  }
}