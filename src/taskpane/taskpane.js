/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Variable global para almacenar las credenciales
let userCredentials = {
  username: null,
  password: null,
  isLoggedIn: false
};

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
        
        // Usar proxy local para evitar problemas de CORS y Mixed Content
        const response = await fetch('/odata/', {
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

// Función auxiliar para hacer peticiones autenticadas
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

// Función para mostrar el modal de descarga
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

// Función para ejecutar la descarga según las opciones seleccionadas
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

    console.log("Tipo de descarga:", downloadType);
    console.log("Límite de registros:", recordLimit);
    console.log("Campos seleccionados:", selectedFields);

    // Cerrar el modal
    document.getElementById("downloadModal").classList.add("hidden");

    // Llamar a la función de descarga con los parámetros
    await download(downloadType, recordLimit, selectedFields);
  } catch (error) {
    console.error("Error en executeDownload:", error);
    showNotification("Error al preparar la descarga", "error");
  }
}

export async function download(downloadType = 'cuentas', recordLimit = '50', selectedFields = []) {
  try {
    // Suspender actualización de pantalla para mejor rendimiento
    await Excel.run(async (context) => {
      const application = context.workbook.application;
      application.suspendScreenUpdatingUntilNextSync();
      
      // Construir la URL según el tipo de descarga
      let endpoint = '';
      switch(downloadType) {
        case 'cuentas':
          endpoint = '/odata/AccountSet';
          break;
        case 'flujos':
          endpoint = '/odata/FlowCodeSet';
          break;
        case 'movimientos':
          endpoint = '/odata/CashFlowSet';
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
          
          // Agregar $filter solo con Status
          params.push("$filter=Status eq 'Actual'");
          
          // Unir todos los parámetros
          if (params.length > 0) {
            endpoint += '?' + params.join('&');
          }
          break;
      }
      
      // Agregar límite de registros para Cuentas y Flujos
      if (downloadType !== 'movimientos' && recordLimit !== 'all') {
        endpoint += (endpoint.includes('?') ? '&' : '?') + `$top=${recordLimit}`;
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
        throw new Error("No se recibieron datos del servidor");
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

      // Función auxiliar para formatear fechas
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

      // Función auxiliar para formatear valores según el tipo de campo
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
      
      // Generar la columna final (A, B, ..., Z, AA, AB, ...)
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

export async function importData() {
  try {
    console.log("Iniciando función importData");
    await Excel.run(async (context) => {
      console.log("Dentro de Excel.run");
      // Suspender actualización de pantalla
      const application = context.workbook.application;
      application.suspendScreenUpdatingUntilNextSync();

      // Validar y obtener datos de origen
      console.log("Obteniendo hoja activa y rango");
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      const range = sheet.getRange("A1:B2");
      range.load(["values", "rowCount", "columnCount"]);
      
      await context.sync();
      console.log("Después de sync, valores obtenidos:", range.values);

      // Validaciones de datos
      if (!range.values || range.values.length < 2) {
        console.log("Error: No hay suficientes datos", range.values);
        throw new Error("No hay suficientes datos para importar");
      }

      if (!range.values[0][0] || !range.values[0][1]) {
        console.log("Error: Faltan encabezados", range.values[0]);
        throw new Error("Los encabezados son requeridos");
      }

      // Validar que los datos no estén vacíos
      if (!range.values[1][0] && !range.values[1][1]) {
        throw new Error("No hay datos para importar");
      }

      const data = {
        title: range.values[1][0] || "",
        body: range.values[1][1] || "",
        userId: 1
      };

      console.log("Preparando datos para enviar:", data);

      // Intentar enviar datos con reintento
      let result;
      let retries = 3;
      while (retries > 0) {
        try {
          console.log(`Intento ${4-retries} de envío de datos`);
          const response = await fetch('https://jsonplaceholder.typicode.com/posts', {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
            },
            body: JSON.stringify(data)
          });

          if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
          }

          result = await response.json();
          break;
        } catch (fetchError) {
          retries--;
          if (retries === 0) throw new Error('Error al enviar datos después de 3 intentos');
          await new Promise(resolve => setTimeout(resolve, 1000));
        }
      }

      // Crear hoja de resultado con nombre único
      let resultSheetName = "Resultado";
      let counter = 1;
      while (true) {
        try {
          context.workbook.worksheets.getItem(resultSheetName);
          resultSheetName = `Resultado_${counter++}`;
        } catch {
          break;
        }
      }

      const resultSheet = context.workbook.worksheets.add(resultSheetName);
      
      // Escribir resultados en un solo bloque
      const resultRange = resultSheet.getRange("A1:C2");
      resultRange.values = [
        ["ID", "Estado", "Fecha"],
        [result.id, "Importado exitosamente", new Date().toLocaleString()]
      ];

      // Formatear la hoja de resultados
      const headerRange = resultRange.getRow(0);
      headerRange.format.fill.color = "#D3D3D3";
      headerRange.format.font.bold = true;
      resultSheet.getUsedRange().format.autofitColumns();
      
      await context.sync();
      showNotification("¡Datos importados exitosamente!", "success");
    });
  } catch (error) {
    console.error("Error específico:", error.message);
    let errorMessage = "Error al importar los datos";
    
    // Mensajes de error más específicos
    if (error.message.includes("suficientes datos")) {
      errorMessage = "No hay suficientes datos para importar. Verifique el rango seleccionado.";
    } else if (error.message.includes("enviar datos")) {
      errorMessage = "Error de conexión al enviar datos. Verifique su conexión a internet.";
    } else if (error.message.includes("encabezados")) {
      errorMessage = "Los encabezados son requeridos. Verifique la estructura de los datos.";
    }
    
    showNotification(errorMessage, "error");
  }
}