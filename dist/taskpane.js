/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./assets/logo-filled.png":
/*!********************************!*\
  !*** ./assets/logo-filled.png ***!
  \********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "assets/logo-filled.png";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "32ab1bef0886af150899.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = (typeof document !== 'undefined' && document.baseURI) || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   download: function() { return /* binding */ download; },
/* harmony export */   importData: function() { return /* binding */ importData; },
/* harmony export */   login: function() { return /* binding */ login; },
/* harmony export */   showDownloadModal: function() { return /* binding */ showDownloadModal; }
/* harmony export */ });
function _toConsumableArray(r) { return _arrayWithoutHoles(r) || _iterableToArray(r) || _unsupportedIterableToArray(r) || _nonIterableSpread(); }
function _nonIterableSpread() { throw new TypeError("Invalid attempt to spread non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); }
function _unsupportedIterableToArray(r, a) { if (r) { if ("string" == typeof r) return _arrayLikeToArray(r, a); var t = {}.toString.call(r).slice(8, -1); return "Object" === t && r.constructor && (t = r.constructor.name), "Map" === t || "Set" === t ? Array.from(r) : "Arguments" === t || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t) ? _arrayLikeToArray(r, a) : void 0; } }
function _iterableToArray(r) { if ("undefined" != typeof Symbol && null != r[Symbol.iterator] || null != r["@@iterator"]) return Array.from(r); }
function _arrayWithoutHoles(r) { if (Array.isArray(r)) return _arrayLikeToArray(r); }
function _arrayLikeToArray(r, a) { (null == a || a > r.length) && (a = r.length); for (var e = 0, n = Array(a); e < a; e++) n[e] = r[e]; return n; }
function _typeof(o) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (o) { return typeof o; } : function (o) { return o && "function" == typeof Symbol && o.constructor === Symbol && o !== Symbol.prototype ? "symbol" : typeof o; }, _typeof(o); }
function ownKeys(e, r) { var t = Object.keys(e); if (Object.getOwnPropertySymbols) { var o = Object.getOwnPropertySymbols(e); r && (o = o.filter(function (r) { return Object.getOwnPropertyDescriptor(e, r).enumerable; })), t.push.apply(t, o); } return t; }
function _objectSpread(e) { for (var r = 1; r < arguments.length; r++) { var t = null != arguments[r] ? arguments[r] : {}; r % 2 ? ownKeys(Object(t), !0).forEach(function (r) { _defineProperty(e, r, t[r]); }) : Object.getOwnPropertyDescriptors ? Object.defineProperties(e, Object.getOwnPropertyDescriptors(t)) : ownKeys(Object(t)).forEach(function (r) { Object.defineProperty(e, r, Object.getOwnPropertyDescriptor(t, r)); }); } return e; }
function _defineProperty(e, r, t) { return (r = _toPropertyKey(r)) in e ? Object.defineProperty(e, r, { value: t, enumerable: !0, configurable: !0, writable: !0 }) : e[r] = t, e; }
function _toPropertyKey(t) { var i = _toPrimitive(t, "string"); return "symbol" == _typeof(i) ? i : i + ""; }
function _toPrimitive(t, r) { if ("object" != _typeof(t) || !t) return t; var e = t[Symbol.toPrimitive]; if (void 0 !== e) { var i = e.call(t, r || "default"); if ("object" != _typeof(i)) return i; throw new TypeError("@@toPrimitive must return a primitive value."); } return ("string" === r ? String : Number)(t); }
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
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
var userCredentials = {
  username: null,
  password: null,
  isLoggedIn: false
};

/**
 * Variable global para almacenar el ID de la cuenta seleccionada
 * @type {string|null}
 */
var selectedAccountId = null;

/**
 * Variable global para almacenar las cuentas descargadas desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
var cachedAccounts = [];

/**
 * Variable global para almacenar los flujos de caja descargados desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
var cachedFlowCodes = [];

/**
 * Variable global para almacenar los códigos presupuestarios descargados desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
var cachedBudgetCodes = [];

/**
 * Variable global para almacenar las divisas descargadas desde OData
 * @type {Array<{Id: string, Code: string}>}
 */
var cachedCurrencies = [];

/**
 * Variable global para almacenar los lugares de cotización descargados desde OData
 * @type {Array<{Id: number, Description: string}>}
 */
var cachedQuotationPlaces = [];

/**
 * Variables globales para almacenar el rango de fechas seleccionado
 * @type {string|null}
 */
var selectedDateFrom = null;
var selectedDateTo = null;

/**
 * Función de inicialización de Office.js
 * Se ejecuta cuando el entorno de Office está listo para interactuar
 * @param {Object} info - Información sobre el host de Office
 */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").classList.add("hidden");
    document.getElementById("app-body").classList.remove("hidden");

    // Agregar event listeners para los botones
    document.getElementById("login").onclick = login;
    document.getElementById("download").onclick = showDownloadModal;
    document.getElementById("import").onclick = importData;

    // Event listener para cambio de tipo de descarga
    document.getElementById("downloadType").onchange = function () {
      var movimientosOptions = document.getElementById("movimientosOptions");
      if (this.value === "movimientos") {
        movimientosOptions.classList.remove("hidden");
        // Cargar cuentas al mostrar opciones de movimientos
        loadAccounts();
      } else {
        movimientosOptions.classList.add("hidden");
      }
    };

    // Event listener para cambio de cuenta seleccionada
    document.getElementById("accountSelect").onchange = function () {
      selectedAccountId = this.value;
      console.log("Cuenta seleccionada:", this.options[this.selectedIndex].text, "(ID:", selectedAccountId, ")");
    };

    // Event listeners para los campos de fecha
    document.getElementById("dateFrom").onchange = function () {
      selectedDateFrom = this.value;
      console.log("Fecha desde seleccionada:", selectedDateFrom);
    };
    document.getElementById("dateTo").onchange = function () {
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
function login() {
  return _login.apply(this, arguments);
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
function _login() {
  _login = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
    var modal, loginSubmitButton, _t3;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.p = _context2.n) {
        case 0:
          _context2.p = 0;
          console.log("Función login iniciada");
          modal = document.getElementById("loginModal");
          if (modal) {
            _context2.n = 1;
            break;
          }
          console.error("Modal no encontrado en el DOM");
          return _context2.a(2);
        case 1:
          console.log("Modal encontrado, removiendo clase hidden");
          modal.classList.remove("hidden");
          modal.style.display = "block"; // Forzar visualización
          loginSubmitButton = document.getElementById("loginSubmit");
          if (loginSubmitButton) {
            _context2.n = 2;
            break;
          }
          console.error("Botón submit no encontrado");
          return _context2.a(2);
        case 2:
          console.log("Configurando evento click del botón submit");
          loginSubmitButton.onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
            var username, password, loadingDiv, errorDiv, submitButton, cancelButton, authString, isDevelopment, baseUrl, response, loginButton, downloadButton, importButton, errorDetails, errorText, _errorDiv, _errorDiv2, errorMsg, _t, _t2;
            return _regenerator().w(function (_context) {
              while (1) switch (_context.p = _context.n) {
                case 0:
                  username = document.getElementById("username").value;
                  password = document.getElementById("password").value;
                  if (!(!username || !password)) {
                    _context.n = 1;
                    break;
                  }
                  console.error("Por favor complete todos los campos");
                  return _context.a(2);
                case 1:
                  // Mostrar spinner y ocultar error previo
                  loadingDiv = document.getElementById("loginLoading");
                  errorDiv = document.getElementById("loginError");
                  submitButton = document.getElementById("loginSubmit");
                  cancelButton = document.getElementById("loginCancel");
                  loadingDiv.classList.remove("hidden");
                  errorDiv.classList.add("hidden");
                  submitButton.disabled = true;
                  cancelButton.disabled = true;
                  _context.p = 2;
                  console.log("Intentando hacer login con:", {
                    username: username
                  });

                  // Crear el header de autenticación básica
                  authString = btoa(username + ':' + password);
                  console.log("Autenticación básica creada");

                  // DESARROLLO: Usar proxy de webpack (/odata)
                  // PRODUCCIÓN: Usar proxy Vercel (https://excel-addin-azurriga.vercel.app)
                  isDevelopment = window.location.hostname === 'localhost';
                  baseUrl = isDevelopment ? '/odata/AccountSet?$top=1' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/AccountSet?$top=1';
                  console.log("Usando proxy ".concat(isDevelopment ? 'WEBPACK' : 'VERCEL', ": ").concat(baseUrl));
                  _context.n = 3;
                  return fetch(baseUrl, {
                    method: 'GET',
                    headers: {
                      'Authorization': "Basic ".concat(authString),
                      'Content-Type': 'application/json'
                    }
                  });
                case 3:
                  response = _context.v;
                  console.log("Respuesta recibida:", response);
                  console.log("Status:", response.status);
                  console.log("Status Text:", response.statusText);

                  // Ocultar spinner
                  loadingDiv.classList.add("hidden");
                  submitButton.disabled = false;
                  cancelButton.disabled = false;
                  if (!response.ok) {
                    _context.n = 4;
                    break;
                  }
                  console.log("Login exitoso");

                  // Guardar las credenciales
                  userCredentials.username = username;
                  userCredentials.password = password;
                  userCredentials.isLoggedIn = true;
                  loginButton = document.getElementById("login");
                  loginButton.innerHTML = "<span class=\"ms-Button-label\">\xA1Bienvenido ".concat(username, "!</span>");
                  loginButton.style.backgroundColor = "#107C10";

                  // Activar los botones de Descargar e Importar
                  downloadButton = document.getElementById("download");
                  importButton = document.getElementById("import");
                  downloadButton.classList.remove("is-disabled");
                  downloadButton.removeAttribute("disabled");
                  importButton.classList.remove("is-disabled");
                  importButton.removeAttribute("disabled");
                  modal.classList.add("hidden");
                  showNotification("¡Sesión iniciada correctamente!", "success");
                  _context.n = 9;
                  break;
                case 4:
                  console.error("Error de autenticación. Status:", response.status);

                  // Leer el cuerpo de la respuesta para más detalles
                  errorDetails = '';
                  _context.p = 5;
                  _context.n = 6;
                  return response.text();
                case 6:
                  errorText = _context.v;
                  errorDetails = " (".concat(response.status, ": ").concat(errorText.substring(0, 100), ")");
                  _context.n = 8;
                  break;
                case 7:
                  _context.p = 7;
                  _t = _context.v;
                  errorDetails = " (C\xF3digo: ".concat(response.status, ")");
                case 8:
                  // Mostrar mensaje de error en el modal
                  _errorDiv = document.getElementById("loginError");
                  _errorDiv.innerHTML = "Usuario o contrase\xF1a incorrectos".concat(errorDetails);
                  _errorDiv.classList.remove("hidden");

                  // Limpiar el mensaje de error después de 5 segundos
                  setTimeout(function () {
                    _errorDiv.classList.add("hidden");
                  }, 5000);
                case 9:
                  _context.n = 11;
                  break;
                case 10:
                  _context.p = 10;
                  _t2 = _context.v;
                  console.error("Error en login (catch):", _t2);
                  console.error("Error message:", _t2.message);
                  console.error("Error stack:", _t2.stack);

                  // Ocultar spinner y reactivar botones
                  loadingDiv.classList.add("hidden");
                  submitButton.disabled = false;
                  cancelButton.disabled = false;
                  _errorDiv2 = document.getElementById("loginError"); // Construir mensaje de error más detallado
                  errorMsg = "Error de conexión: ";
                  if (_t2.message.includes('Failed to fetch')) {
                    errorMsg += "No se puede conectar al servidor. Verifique:\n1. La URL del servidor\n2. Que el servidor esté en ejecución\n3. Configuración de CORS en el servidor";
                  } else if (_t2.message.includes('NetworkError')) {
                    errorMsg += "Error de red. Verifique su conexión a Internet.";
                  } else {
                    errorMsg += _t2.message;
                  }
                  _errorDiv2.innerHTML = errorMsg.replace(/\n/g, '<br>');
                  _errorDiv2.classList.remove("hidden");

                  // Limpiar el mensaje de error después de 7 segundos
                  setTimeout(function () {
                    _errorDiv2.classList.add("hidden");
                  }, 7000);
                case 11:
                  return _context.a(2);
              }
            }, _callee, null, [[5, 7], [2, 10]]);
          }));
          document.getElementById("loginCancel").onclick = function () {
            modal.classList.add("hidden");
          };
          window.onclick = function (event) {
            if (event.target === modal) {
              modal.classList.add("hidden");
            }
          };
          _context2.n = 4;
          break;
        case 3:
          _context2.p = 3;
          _t3 = _context2.v;
          console.error("Error:", _t3);
        case 4:
          return _context2.a(2);
      }
    }, _callee2, null, [[0, 3]]);
  }));
  return _login.apply(this, arguments);
}
function showNotification(message) {
  var type = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : 'success';
  var popup = document.getElementById('notificationPopup');
  var messageEl = document.getElementById('notificationMessage');

  // Establecer el mensaje
  messageEl.textContent = message;

  // Aplicar clase de estilo según el tipo
  popup.classList.remove('success', 'error');
  popup.classList.add(type);

  // Mostrar el popup
  popup.classList.remove('hidden');

  // Ocultar después de 3 segundos
  setTimeout(function () {
    popup.classList.add('hidden');
  }, 3000);
}

/**
 * Muestra un panel modal con errores detallados
 * @param {string} message - Mensaje detallado de errores
 */
function showErrorDetails(message) {
  var panel = document.getElementById('errorDetailsPanel');
  var messageEl = document.getElementById('errorDetailsMessage');

  // Establecer el mensaje
  messageEl.textContent = message;

  // Mostrar el panel
  panel.classList.remove('hidden');
}

/**
 * Oculta el panel de errores detallados
 */
function hideErrorDetails() {
  var panel = document.getElementById('errorDetailsPanel');
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
function loadAccounts() {
  return _loadAccounts.apply(this, arguments);
}
/**
 * Carga los flujos de caja desde el servidor
 * 
 * Consulta: odata/FlowCodeSet?$select=Code,Id
 * 
 * @async
 */
function _loadAccounts() {
  _loadAccounts = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var accountSelect, loadingOption, isDevelopment, VERCEL_PROXY, separator, ampersand, endpoint, response, data, noDataOption, _accountSelect, _t4;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.p = _context3.n) {
        case 0:
          _context3.p = 0;
          accountSelect = document.getElementById("accountSelect"); // Limpiar opciones existentes (excepto la primera "Todas las cuentas")
          accountSelect.innerHTML = '<option value="">Todas las cuentas</option>';

          // Mostrar indicador de carga
          loadingOption = document.createElement('option');
          loadingOption.value = '';
          loadingOption.textContent = 'Cargando cuentas...';
          loadingOption.disabled = true;
          accountSelect.appendChild(loadingOption);

          // Determinar el proxy correcto
          isDevelopment = window.location.hostname === 'localhost';
          VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/'; // Construir URL con filtro y select
          separator = isDevelopment ? '?' : '%3F';
          ampersand = isDevelopment ? '&' : '%26';
          endpoint = "".concat(VERCEL_PROXY, "AccountSet").concat(separator, "$filter=Active eq true").concat(ampersand, "$select=Code,Id");
          console.log("Cargando cuentas desde:", endpoint);

          // Obtener las cuentas con header especial para números grandes
          _context3.n = 1;
          return authenticatedFetch(endpoint, {
            headers: {
              'Accept': 'application/json;IEEE754Compatible=true'
            }
          });
        case 1:
          response = _context3.v;
          if (response.ok) {
            _context3.n = 2;
            break;
          }
          throw new Error("Error al cargar cuentas: ".concat(response.status));
        case 2:
          _context3.n = 3;
          return response.json();
        case 3:
          data = _context3.v;
          console.log("Cuentas recibidas:", data);

          // Limpiar el indicador de carga
          accountSelect.innerHTML = '<option value="">Todas las cuentas</option>';

          // Verificar que tengamos datos
          if (data && data.value && data.value.length > 0) {
            // Almacenar las cuentas en el caché global
            cachedAccounts = data.value;

            // Agregar cada cuenta al combo
            data.value.forEach(function (account) {
              var option = document.createElement('option');
              option.value = account.Id; // Valor interno: ID
              option.textContent = account.Code; // Texto visible: Code
              accountSelect.appendChild(option);
            });
            console.log("".concat(data.value.length, " cuentas cargadas correctamente"));
          } else {
            // No hay cuentas activas
            noDataOption = document.createElement('option');
            noDataOption.value = '';
            noDataOption.textContent = 'No hay cuentas activas disponibles';
            noDataOption.disabled = true;
            accountSelect.appendChild(noDataOption);
          }
          _context3.n = 5;
          break;
        case 4:
          _context3.p = 4;
          _t4 = _context3.v;
          console.error("Error cargando cuentas:", _t4);

          // Mostrar error en el combo
          _accountSelect = document.getElementById("accountSelect");
          _accountSelect.innerHTML = '<option value="">Error al cargar cuentas</option>';

          // Mostrar notificación al usuario
          showNotification("Error al cargar las cuentas: " + _t4.message, "error");
        case 5:
          return _context3.a(2);
      }
    }, _callee3, null, [[0, 4]]);
  }));
  return _loadAccounts.apply(this, arguments);
}
function loadFlowCodes() {
  return _loadFlowCodes.apply(this, arguments);
}
/**
 * Carga los códigos presupuestarios desde el servidor
 * 
 * Consulta: odata/BudgetCodeSet?$select=Code,Id
 * 
 * @async
 */
function _loadFlowCodes() {
  _loadFlowCodes = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
    var isDevelopment, VERCEL_PROXY, separator, endpoint, response, data, _t5;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.p = _context4.n) {
        case 0:
          _context4.p = 0;
          console.log("Iniciando carga de flujos de caja...");
          isDevelopment = window.location.hostname === 'localhost';
          VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
          separator = isDevelopment ? '?' : '%3F';
          endpoint = "".concat(VERCEL_PROXY, "FlowCodeSet").concat(separator, "$select=Code,Id");
          console.log("Cargando flujos desde:", endpoint);
          _context4.n = 1;
          return authenticatedFetch(endpoint, {
            headers: {
              'Accept': 'application/json;IEEE754Compatible=true'
            }
          });
        case 1:
          response = _context4.v;
          if (response.ok) {
            _context4.n = 2;
            break;
          }
          throw new Error("Error al cargar flujos: ".concat(response.status));
        case 2:
          _context4.n = 3;
          return response.json();
        case 3:
          data = _context4.v;
          console.log("Flujos recibidos:", data);
          if (data && data.value && data.value.length > 0) {
            cachedFlowCodes = data.value;
            console.log("".concat(data.value.length, " flujos de caja cargados correctamente"));
          } else {
            console.warn("No hay flujos de caja disponibles");
          }
          _context4.n = 5;
          break;
        case 4:
          _context4.p = 4;
          _t5 = _context4.v;
          console.error("Error cargando flujos de caja:", _t5);
          showNotification("Error al cargar flujos de caja: " + _t5.message, "error");
        case 5:
          return _context4.a(2);
      }
    }, _callee4, null, [[0, 4]]);
  }));
  return _loadFlowCodes.apply(this, arguments);
}
function loadBudgetCodes() {
  return _loadBudgetCodes.apply(this, arguments);
}
/**
 * Carga las divisas desde el servidor
 * 
 * Consulta: odata/CurrencySet (sin $select para obtener todos los campos)
 * 
 * @async
 */
function _loadBudgetCodes() {
  _loadBudgetCodes = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5() {
    var isDevelopment, VERCEL_PROXY, separator, endpoint, response, data, _t6;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          _context5.p = 0;
          console.log("Iniciando carga de códigos presupuestarios...");
          isDevelopment = window.location.hostname === 'localhost';
          VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
          separator = isDevelopment ? '?' : '%3F';
          endpoint = "".concat(VERCEL_PROXY, "BudgetCodeSet").concat(separator, "$select=Code,Id");
          console.log("Cargando códigos presupuestarios desde:", endpoint);
          _context5.n = 1;
          return authenticatedFetch(endpoint, {
            headers: {
              'Accept': 'application/json;IEEE754Compatible=true'
            }
          });
        case 1:
          response = _context5.v;
          if (response.ok) {
            _context5.n = 2;
            break;
          }
          throw new Error("Error al cargar c\xF3digos presupuestarios: ".concat(response.status));
        case 2:
          _context5.n = 3;
          return response.json();
        case 3:
          data = _context5.v;
          console.log("Códigos presupuestarios recibidos:", data);
          if (data && data.value && data.value.length > 0) {
            cachedBudgetCodes = data.value;
            console.log("".concat(data.value.length, " c\xF3digos presupuestarios cargados correctamente"));
          } else {
            console.warn("No hay códigos presupuestarios disponibles");
          }
          _context5.n = 5;
          break;
        case 4:
          _context5.p = 4;
          _t6 = _context5.v;
          console.error("Error cargando códigos presupuestarios:", _t6);
          showNotification("Error al cargar códigos presupuestarios: " + _t6.message, "error");
        case 5:
          return _context5.a(2);
      }
    }, _callee5, null, [[0, 4]]);
  }));
  return _loadBudgetCodes.apply(this, arguments);
}
function loadCurrencies() {
  return _loadCurrencies.apply(this, arguments);
}
/**
 * Carga los lugares de cotización desde el servidor
 * 
 * Consulta: odata/QuotationPlaceSet (sin $select para obtener todos los campos)
 * 
 * @async
 */
function _loadCurrencies() {
  _loadCurrencies = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6() {
    var isDevelopment, VERCEL_PROXY, endpoint, response, data, _t7;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.p = _context6.n) {
        case 0:
          _context6.p = 0;
          console.log("Iniciando carga de divisas...");
          isDevelopment = window.location.hostname === 'localhost';
          VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/'; // No usar $select para obtener todos los campos disponibles
          endpoint = "".concat(VERCEL_PROXY, "CurrencySet");
          console.log("Cargando divisas desde:", endpoint);
          _context6.n = 1;
          return authenticatedFetch(endpoint, {
            headers: {
              'Accept': 'application/json;IEEE754Compatible=true'
            }
          });
        case 1:
          response = _context6.v;
          if (response.ok) {
            _context6.n = 2;
            break;
          }
          throw new Error("Error al cargar divisas: ".concat(response.status));
        case 2:
          _context6.n = 3;
          return response.json();
        case 3:
          data = _context6.v;
          console.log("Divisas recibidas:", data);
          if (data && data.value && data.value.length > 0) {
            // Mapear los datos según la estructura real de CurrencySet
            // Asumiendo que tiene Id como primary key
            cachedCurrencies = data.value.map(function (curr) {
              return {
                Id: curr.Id || curr.Code || curr.id,
                Code: curr.Code || curr.Id || curr.id
              };
            });
            console.log("".concat(cachedCurrencies.length, " divisas cargadas correctamente"));
            console.log("Ejemplo de divisa:", cachedCurrencies[0]);
          } else {
            console.warn("No hay divisas disponibles");
          }
          _context6.n = 5;
          break;
        case 4:
          _context6.p = 4;
          _t7 = _context6.v;
          console.error("Error cargando divisas:", _t7);
          showNotification("Error al cargar divisas: " + _t7.message, "error");
        case 5:
          return _context6.a(2);
      }
    }, _callee6, null, [[0, 4]]);
  }));
  return _loadCurrencies.apply(this, arguments);
}
function loadQuotationPlaces() {
  return _loadQuotationPlaces.apply(this, arguments);
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
function _loadQuotationPlaces() {
  _loadQuotationPlaces = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7() {
    var isDevelopment, VERCEL_PROXY, endpoint, response, data, _t8;
    return _regenerator().w(function (_context7) {
      while (1) switch (_context7.p = _context7.n) {
        case 0:
          _context7.p = 0;
          console.log("Iniciando carga de lugares de cotización...");
          isDevelopment = window.location.hostname === 'localhost';
          VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/'; // No usar $select para obtener todos los campos disponibles
          endpoint = "".concat(VERCEL_PROXY, "QuotationPlaceSet");
          console.log("Cargando lugares de cotización desde:", endpoint);
          _context7.n = 1;
          return authenticatedFetch(endpoint, {
            headers: {
              'Accept': 'application/json;IEEE754Compatible=true'
            }
          });
        case 1:
          response = _context7.v;
          if (response.ok) {
            _context7.n = 2;
            break;
          }
          throw new Error("Error al cargar lugares de cotizaci\xF3n: ".concat(response.status));
        case 2:
          _context7.n = 3;
          return response.json();
        case 3:
          data = _context7.v;
          console.log("Lugares de cotización recibidos:", data);
          if (data && data.value && data.value.length > 0) {
            // Mapear los datos según la estructura real de QuotationPlaceSet
            cachedQuotationPlaces = data.value.map(function (qp) {
              return {
                Id: qp.Id || qp.id,
                Description: qp.Description || qp.Name || qp.description || qp.name || "Cotizaci\xF3n ".concat(qp.Id)
              };
            });
            console.log("".concat(cachedQuotationPlaces.length, " lugares de cotizaci\xF3n cargados correctamente"));
            console.log("Ejemplo de lugar de cotización:", cachedQuotationPlaces[0]);
          } else {
            console.warn("No hay lugares de cotización disponibles");
          }
          _context7.n = 5;
          break;
        case 4:
          _context7.p = 4;
          _t8 = _context7.v;
          console.error("Error cargando lugares de cotización:", _t8);
          showNotification("Error al cargar lugares de cotización: " + _t8.message, "error");
        case 5:
          return _context7.a(2);
      }
    }, _callee7, null, [[0, 4]]);
  }));
  return _loadQuotationPlaces.apply(this, arguments);
}
function authenticatedFetch(_x) {
  return _authenticatedFetch.apply(this, arguments);
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
function _authenticatedFetch() {
  _authenticatedFetch = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8(url) {
    var options,
      defaultHeaders,
      mergedOptions,
      _args8 = arguments;
    return _regenerator().w(function (_context8) {
      while (1) switch (_context8.n) {
        case 0:
          options = _args8.length > 1 && _args8[1] !== undefined ? _args8[1] : {};
          if (userCredentials.isLoggedIn) {
            _context8.n = 1;
            break;
          }
          throw new Error("Debe iniciar sesión primero");
        case 1:
          defaultHeaders = {
            'Content-Type': 'application/json; charset=utf-8',
            'Accept': 'application/json; charset=utf-8',
            'Authorization': "Basic ".concat(btoa(userCredentials.username + ':' + userCredentials.password))
          }; // Mezclar headers personalizados con los predeterminados
          mergedOptions = _objectSpread(_objectSpread({}, options), {}, {
            headers: _objectSpread(_objectSpread({}, defaultHeaders), options.headers || {})
          });
          return _context8.a(2, fetch(url, mergedOptions));
      }
    }, _callee8);
  }));
  return _authenticatedFetch.apply(this, arguments);
}
function showDownloadModal() {
  return _showDownloadModal.apply(this, arguments);
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
function _showDownloadModal() {
  _showDownloadModal = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0() {
    var modal, _t9;
    return _regenerator().w(function (_context0) {
      while (1) switch (_context0.p = _context0.n) {
        case 0:
          _context0.p = 0;
          if (userCredentials.isLoggedIn) {
            _context0.n = 1;
            break;
          }
          showNotification("Debe iniciar sesión primero", "error");
          return _context0.a(2);
        case 1:
          modal = document.getElementById("downloadModal");
          modal.classList.remove("hidden");
          modal.style.display = "block";

          // Configurar botón de submit
          document.getElementById("downloadSubmit").onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9() {
            return _regenerator().w(function (_context9) {
              while (1) switch (_context9.n) {
                case 0:
                  _context9.n = 1;
                  return executeDownload();
                case 1:
                  return _context9.a(2);
              }
            }, _callee9);
          }));

          // Configurar botón de cancelar
          document.getElementById("downloadCancel").onclick = function () {
            modal.classList.add("hidden");
          };

          // Cerrar modal al hacer clic fuera
          window.onclick = function (event) {
            if (event.target === modal) {
              modal.classList.add("hidden");
            }
          };
          _context0.n = 3;
          break;
        case 2:
          _context0.p = 2;
          _t9 = _context0.v;
          console.error("Error al abrir modal de descarga:", _t9);
          showNotification("Error al abrir el modal de descarga", "error");
        case 3:
          return _context0.a(2);
      }
    }, _callee0, null, [[0, 2]]);
  }));
  return _showDownloadModal.apply(this, arguments);
}
function executeDownload() {
  return _executeDownload.apply(this, arguments);
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
function _executeDownload() {
  _executeDownload = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee1() {
    var downloadType, recordLimit, selectedFields, checkboxes, _t0;
    return _regenerator().w(function (_context1) {
      while (1) switch (_context1.p = _context1.n) {
        case 0:
          _context1.p = 0;
          downloadType = document.getElementById("downloadType").value;
          recordLimit = document.getElementById("recordLimit").value; // Recoger campos seleccionados para Movimientos
          selectedFields = [];
          if (!(downloadType === "movimientos")) {
            _context1.n = 1;
            break;
          }
          checkboxes = document.querySelectorAll('#movimientosOptions input[type="checkbox"]:checked');
          selectedFields = Array.from(checkboxes).map(function (cb) {
            return cb.value;
          });
          if (!(selectedFields.length === 0)) {
            _context1.n = 1;
            break;
          }
          showNotification("Debe seleccionar al menos un campo", "error");
          return _context1.a(2);
        case 1:
          console.log("Descarga:", downloadType, "| Registros:", recordLimit, "| Cuenta:", selectedAccountId || "Todas", "| Desde:", selectedDateFrom || "N/A", "| Hasta:", selectedDateTo || "N/A");

          // Cerrar el modal
          document.getElementById("downloadModal").classList.add("hidden");

          // Llamar a la función de descarga con los parámetros
          _context1.n = 2;
          return download(downloadType, recordLimit, selectedFields);
        case 2:
          _context1.n = 4;
          break;
        case 3:
          _context1.p = 3;
          _t0 = _context1.v;
          console.error("Error en executeDownload:", _t0);
          showNotification("Error al preparar la descarga", "error");
        case 4:
          return _context1.a(2);
      }
    }, _callee1, null, [[0, 3]]);
  }));
  return _executeDownload.apply(this, arguments);
}
function download() {
  return _download.apply(this, arguments);
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
function _download() {
  _download = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee11() {
    var downloadType,
      recordLimit,
      selectedFields,
      errorMessage,
      _args11 = arguments,
      _t14;
    return _regenerator().w(function (_context11) {
      while (1) switch (_context11.p = _context11.n) {
        case 0:
          downloadType = _args11.length > 0 && _args11[0] !== undefined ? _args11[0] : 'cuentas';
          recordLimit = _args11.length > 1 && _args11[1] !== undefined ? _args11[1] : '50';
          selectedFields = _args11.length > 2 && _args11[2] !== undefined ? _args11[2] : [];
          _context11.p = 1;
          _context11.n = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref3 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee10(context) {
              var application, isDevelopment, VERCEL_PROXY, endpoint, budgetParams, questionMark, ampersand, params, expandParam, filterConditions, filterParam, _questionMark, _ampersand, hasParams, separator, response, retries, data, records, sheetName, existingSheet, sheet, sheet1, formatDate, formatValue, headers, values, allFields, numRows, numCols, getColumnLetter, endColumn, range, headerRange, dateFields, idColIndex, idColLetter, idRange, _t1, _t10, _t11, _t12, _t13;
              return _regenerator().w(function (_context10) {
                while (1) switch (_context10.p = _context10.n) {
                  case 0:
                    application = context.workbook.application;
                    application.suspendScreenUpdatingUntilNextSync();

                    // DESARROLLO: Usar proxy de webpack (/odata)
                    // PRODUCCIÓN: Usar proxy Vercel (https://excel-addin-azurriga.vercel.app)
                    isDevelopment = window.location.hostname === 'localhost';
                    VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
                    console.log("Download usando proxy ".concat(isDevelopment ? 'WEBPACK' : 'VERCEL'));
                    endpoint = '';
                    _t1 = downloadType;
                    _context10.n = _t1 === 'cuentas' ? 1 : _t1 === 'flujos' ? 2 : _t1 === 'codigos-presupuestarios' ? 3 : _t1 === 'divisas' ? 4 : _t1 === 'cotizacion' ? 5 : _t1 === 'movimientos' ? 6 : 7;
                    break;
                  case 1:
                    endpoint = "".concat(VERCEL_PROXY, "AccountSet");
                    return _context10.a(3, 7);
                  case 2:
                    endpoint = "".concat(VERCEL_PROXY, "FlowCodeSet");
                    return _context10.a(3, 7);
                  case 3:
                    endpoint = "".concat(VERCEL_PROXY, "BudgetCodeSet");
                    // Especificar los campos solicitados: Code, Id, Description
                    budgetParams = ['$select=Code,Id,Description'];
                    if (recordLimit !== 'all') {
                      budgetParams.push("$top=".concat(recordLimit));
                    }
                    if (budgetParams.length > 0) {
                      questionMark = isDevelopment ? '?' : '%3F';
                      ampersand = isDevelopment ? '&' : '%26';
                      endpoint += questionMark + budgetParams.join(ampersand);
                    }
                    return _context10.a(3, 7);
                  case 4:
                    endpoint = "".concat(VERCEL_PROXY, "CurrencySet");
                    return _context10.a(3, 7);
                  case 5:
                    endpoint = "".concat(VERCEL_PROXY, "QuotationPlaceSet");
                    return _context10.a(3, 7);
                  case 6:
                    endpoint = "".concat(VERCEL_PROXY, "CashFlowSet");
                    // Construir la URL completa con $select, $expand y $filter
                    params = []; // Agregar límite de registros si no es "all"
                    if (recordLimit !== 'all') {
                      params.push("$top=".concat(recordLimit));
                    }

                    // Agregar $select con los campos seleccionados
                    if (selectedFields.length > 0) {
                      params.push("$select=".concat(selectedFields.join(',')));
                    }

                    // Agregar $expand (siempre se incluye para Movimientos)
                    expandParam = '$expand=FlowCode($select=Code),BudgetCode($select=Code),Account($expand=Master($select=Code);$select=Id),TrnCurrency($select=Id)';
                    params.push(expandParam);

                    // Construir $filter con Status y opcionalmente con AccountId y fechas
                    filterConditions = ["Status eq 'Actual'"]; // Agregar filtro por cuenta si hay una seleccionada
                    if (selectedAccountId) {
                      filterConditions.push("Account/Id eq ".concat(selectedAccountId));
                    }

                    // Agregar filtro por fecha inicio si hay una seleccionada
                    if (selectedDateFrom) {
                      filterConditions.push("ValueDate ge ".concat(selectedDateFrom));
                    }

                    // Agregar filtro por fecha fin si hay una seleccionada
                    if (selectedDateTo) {
                      filterConditions.push("ValueDate le ".concat(selectedDateTo));
                    }

                    // Unir las condiciones del filtro con 'and'
                    filterParam = "$filter=".concat(filterConditions.join(' and '));
                    params.push(filterParam);

                    // Unir todos los parámetros
                    // En desarrollo (webpack): usar ? y & normales
                    // En producción (Vercel): codificar como %3F y %26
                    if (params.length > 0) {
                      _questionMark = isDevelopment ? '?' : '%3F';
                      _ampersand = isDevelopment ? '&' : '%26';
                      endpoint += _questionMark + params.join(_ampersand);
                    }
                    return _context10.a(3, 7);
                  case 7:
                    // Agregar límite de registros para Cuentas y Flujos (codigos-presupuestarios, divisas y cotizacion ya lo gestionan dentro del switch)
                    if (downloadType !== 'movimientos' && downloadType !== 'codigos-presupuestarios' && downloadType !== 'divisas' && downloadType !== 'cotizacion' && recordLimit !== 'all') {
                      hasParams = isDevelopment ? endpoint.includes('?') : endpoint.includes('%3F');
                      separator = hasParams ? isDevelopment ? '&' : '%26' : isDevelopment ? '?' : '%3F';
                      endpoint += separator + "$top=".concat(recordLimit);
                    }
                    console.log("Descargando desde:", endpoint);
                    console.log("Usuario autenticado:", userCredentials.username);

                    // Intentar obtener los datos con autenticación
                    retries = 3;
                  case 8:
                    if (!(retries > 0)) {
                      _context10.n = 15;
                      break;
                    }
                    _context10.p = 9;
                    _context10.n = 10;
                    return authenticatedFetch(endpoint);
                  case 10:
                    response = _context10.v;
                    console.log("Respuesta recibida. Status:", response.status);
                    if (!response.ok) {
                      _context10.n = 11;
                      break;
                    }
                    return _context10.a(3, 15);
                  case 11:
                    _context10.n = 14;
                    break;
                  case 12:
                    _context10.p = 12;
                    _t10 = _context10.v;
                    console.error("Error en intento de fetch:", _t10);
                    retries--;
                    if (!(retries === 0)) {
                      _context10.n = 13;
                      break;
                    }
                    throw new Error('Error al obtener datos después de 3 intentos');
                  case 13:
                    _context10.n = 14;
                    return new Promise(function (resolve) {
                      return setTimeout(resolve, 1000);
                    });
                  case 14:
                    _context10.n = 8;
                    break;
                  case 15:
                    _context10.n = 16;
                    return response.json();
                  case 16:
                    data = _context10.v;
                    console.log("Datos recibidos:", data);

                    // Verificar que tengamos datos
                    if (!(!data || !data.value || data.value.length === 0)) {
                      _context10.n = 17;
                      break;
                    }
                    showNotification("No se encontraron registros con los filtros seleccionados", "error");
                    return _context10.a(2);
                  case 17:
                    records = data.value; // OData devuelve los datos en data.value
                    // Determinar el nombre de la hoja según el tipo de descarga
                    sheetName = '';
                    _t11 = downloadType;
                    _context10.n = _t11 === 'cuentas' ? 18 : _t11 === 'flujos' ? 19 : _t11 === 'codigos-presupuestarios' ? 20 : _t11 === 'divisas' ? 21 : _t11 === 'cotizacion' ? 22 : _t11 === 'movimientos' ? 23 : 24;
                    break;
                  case 18:
                    sheetName = 'Accounts';
                    return _context10.a(3, 25);
                  case 19:
                    sheetName = 'Flujos';
                    return _context10.a(3, 25);
                  case 20:
                    sheetName = 'Codigos Presupuestarios';
                    return _context10.a(3, 25);
                  case 21:
                    sheetName = 'Divisas';
                    return _context10.a(3, 25);
                  case 22:
                    sheetName = 'Cotizacion';
                    return _context10.a(3, 25);
                  case 23:
                    sheetName = 'Movimientos';
                    return _context10.a(3, 25);
                  case 24:
                    sheetName = downloadType;
                  case 25:
                    _context10.p = 25;
                    existingSheet = context.workbook.worksheets.getItem(sheetName);
                    existingSheet.delete();
                    _context10.n = 26;
                    return context.sync();
                  case 26:
                    console.log("Hoja existente '".concat(sheetName, "' eliminada"));
                    _context10.n = 28;
                    break;
                  case 27:
                    _context10.p = 27;
                    _t12 = _context10.v;
                    // La hoja no existe, no hay problema
                    console.log("La hoja '".concat(sheetName, "' no existe, se crear\xE1 una nueva"));
                  case 28:
                    // Crear la hoja
                    sheet = context.workbook.worksheets.add(sheetName);
                    sheet.load(["protection", "name"]);
                    _context10.n = 29;
                    return context.sync();
                  case 29:
                    if (!sheet.protection.protected) {
                      _context10.n = 30;
                      break;
                    }
                    throw new Error("La hoja está protegida. No se pueden escribir datos.");
                  case 30:
                    console.log("Hoja creada: ".concat(sheetName));

                    // Eliminar Sheet1 si existe (solo la primera vez)
                    _context10.p = 31;
                    sheet1 = context.workbook.worksheets.getItem("Sheet1");
                    sheet1.delete();
                    _context10.n = 32;
                    return context.sync();
                  case 32:
                    console.log("Hoja Sheet1 eliminada");
                    _context10.n = 34;
                    break;
                  case 33:
                    _context10.p = 33;
                    _t13 = _context10.v;
                    // Sheet1 no existe o ya fue eliminada, continuar normalmente
                    console.log("Sheet1 no existe o ya fue eliminada");
                  case 34:
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
                    formatDate = function formatDate(dateString, fieldName) {
                      // Verificar si el valor es nulo, undefined o string vacío
                      if (!dateString || dateString === '' || dateString === null || dateString === undefined) {
                        console.log("Campo ".concat(fieldName, ": valor vac\xEDo"));
                        return '';
                      }
                      console.log("Formateando ".concat(fieldName, ":"), dateString, 'Tipo:', _typeof(dateString));
                      try {
                        var date = new Date(dateString);

                        // Verificar si la fecha es válida
                        if (isNaN(date.getTime())) {
                          console.warn("Fecha inv\xE1lida en ".concat(fieldName, ":"), dateString);
                          return '';
                        }

                        // Convertir a número de serie de Excel
                        // Excel cuenta los días desde 1/1/1900 (pero tiene un bug del año 1900)
                        // JavaScript Date empieza desde 1/1/1970
                        // Fórmula: (fecha en ms - fecha base) / ms por día + offset de Excel
                        var excelEpoch = new Date(1899, 11, 30); // 30 de diciembre de 1899
                        var msPerDay = 24 * 60 * 60 * 1000;
                        var excelSerialDate = (date.getTime() - excelEpoch.getTime()) / msPerDay;
                        console.log("".concat(fieldName, " - Excel serial:"), excelSerialDate);
                        return excelSerialDate;
                      } catch (e) {
                        console.error("Error al formatear fecha ".concat(fieldName, ":"), dateString, e);
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
                    formatValue = function formatValue(fieldName, value) {
                      // No mostrar @odata.etag
                      if (fieldName === '@odata.etag') return null;

                      // Formatear booleanos
                      if (fieldName === 'Active' || fieldName === 'HasWarnings' || fieldName === 'IsInterco') {
                        return value === true ? 'true' : value === false ? 'false' : '';
                      }

                      // Formatear fechas
                      if (fieldName === 'CreationDateTime' || fieldName === 'ModificationDateTime' || fieldName === 'BankClosingDate' || fieldName === 'CloseDate' || fieldName === 'ValueDate' || fieldName === 'TrnDate') {
                        return formatDate(value, fieldName);
                      }

                      // Convertir Id a String explícitamente con apóstrofe para forzar formato texto
                      if (fieldName === 'Id') {
                        // Agregar un espacio de ancho cero al inicio para forzar que Excel lo trate como texto
                        return value !== undefined && value !== null ? "'" + String(value) : '';
                      }

                      // Para el resto de campos, devolver tal cual
                      return value !== undefined && value !== null ? value : '';
                    }; // Preparar encabezados y datos según el tipo de descarga
                    headers = [];
                    values = [];
                    if (downloadType === 'movimientos' && selectedFields.length > 0) {
                      // Usar solo los campos seleccionados
                      headers = selectedFields;
                      values = records.map(function (record) {
                        return selectedFields.map(function (field) {
                          return formatValue(field, record[field]);
                        });
                      });
                    } else {
                      // Obtener todos los campos del primer registro, excluyendo @odata.etag
                      allFields = Object.keys(records[0]).filter(function (key) {
                        return key !== '@odata.etag';
                      });
                      headers = allFields;
                      values = records.map(function (record) {
                        return allFields.map(function (field) {
                          return formatValue(field, record[field]);
                        });
                      });
                    }

                    // Calcular el rango necesario
                    numRows = values.length + 1; // +1 para la fila de encabezados
                    numCols = headers.length;
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
                    getColumnLetter = function getColumnLetter(colIndex) {
                      var letter = '';
                      while (colIndex >= 0) {
                        letter = String.fromCharCode(colIndex % 26 + 65) + letter;
                        colIndex = Math.floor(colIndex / 26) - 1;
                      }
                      return letter;
                    };
                    endColumn = getColumnLetter(numCols - 1); // Escribir datos en un solo bloque para mejor rendimiento
                    range = sheet.getRange("A1:".concat(endColumn).concat(numRows));
                    range.values = [headers].concat(_toConsumableArray(values));

                    // Aplicar formato en una sola operación
                    headerRange = range.getRow(0);
                    headerRange.format.fill.color = "#4472C4";
                    headerRange.format.font.bold = true;
                    headerRange.format.font.color = "#FFFFFF";

                    // Aplicar formato de fecha a las columnas de fecha
                    dateFields = ['CreationDateTime', 'ModificationDateTime', 'BankClosingDate', 'CloseDate', 'ValueDate', 'TrnDate'];
                    dateFields.forEach(function (dateField) {
                      var colIndex = headers.indexOf(dateField);
                      if (colIndex >= 0) {
                        var colLetter = getColumnLetter(colIndex);
                        var dateRange = sheet.getRange("".concat(colLetter, "2:").concat(colLetter).concat(numRows));
                        dateRange.numberFormat = [["DD/MM/YYYY"]];
                        console.log("Formato de fecha aplicado a columna ".concat(colLetter, " (").concat(dateField, ")"));
                      }
                    });

                    // Aplicar formato de texto a la columna Id para evitar notación científica
                    idColIndex = headers.indexOf('Id');
                    if (idColIndex >= 0) {
                      idColLetter = getColumnLetter(idColIndex);
                      idRange = sheet.getRange("".concat(idColLetter, "2:").concat(idColLetter).concat(numRows));
                      idRange.numberFormat = [["@"]]; // @ significa formato texto en Excel
                      console.log("Formato de texto aplicado a columna ".concat(idColLetter, " (Id)"));
                    }

                    // Autoajustar columnas
                    range.format.autofitColumns();

                    // Activar la hoja para que el foco se quede en ella
                    sheet.activate();
                    _context10.n = 35;
                    return context.sync();
                  case 35:
                    showNotification("\xA1".concat(records.length, " ").concat(downloadType, " descargados exitosamente!"), "success");
                  case 36:
                    return _context10.a(2);
                }
              }, _callee10, null, [[31, 33], [25, 27], [9, 12]]);
            }));
            return function (_x3) {
              return _ref3.apply(this, arguments);
            };
          }());
        case 2:
          _context11.n = 4;
          break;
        case 3:
          _context11.p = 3;
          _t14 = _context11.v;
          console.error("Error específico:", _t14.message);
          errorMessage = "Error al descargar los datos"; // Mensajes de error más específicos
          if (_t14.message.includes("protegida")) {
            errorMessage = "La hoja está protegida. Desproteja la hoja e intente nuevamente.";
          } else if (_t14.message.includes("obtener datos")) {
            errorMessage = "Error de conexión. Verifique su conexión a internet.";
          }
          showNotification(errorMessage, "error");
        case 4:
          return _context11.a(2);
      }
    }, _callee11, null, [[1, 3]]);
  }));
  return _download.apply(this, arguments);
}
function importData() {
  return _importData.apply(this, arguments);
}

/**
 * Crea una hoja de Excel con las cabeceras según el tipo seleccionado
 * 
 * @async
 * @throws {Error} Si no se seleccionó un tipo o hay problemas al crear la hoja
 */
function _importData() {
  _importData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee14() {
    var modal, _t15;
    return _regenerator().w(function (_context14) {
      while (1) switch (_context14.p = _context14.n) {
        case 0:
          _context14.p = 0;
          if (userCredentials.isLoggedIn) {
            _context14.n = 1;
            break;
          }
          showNotification("Debe iniciar sesión primero", "error");
          return _context14.a(2);
        case 1:
          modal = document.getElementById("importModal");
          modal.classList.remove("hidden");
          modal.style.display = "block";

          // Configurar botón de crear hoja
          document.getElementById("importCreateSheet").onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee12() {
            return _regenerator().w(function (_context12) {
              while (1) switch (_context12.n) {
                case 0:
                  _context12.n = 1;
                  return executeCreateSheet();
                case 1:
                  return _context12.a(2);
              }
            }, _callee12);
          }));

          // Configurar botón de submit
          document.getElementById("importSubmit").onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee13() {
            return _regenerator().w(function (_context13) {
              while (1) switch (_context13.n) {
                case 0:
                  _context13.n = 1;
                  return executeImport();
                case 1:
                  return _context13.a(2);
              }
            }, _callee13);
          }));

          // Configurar botón de cancelar
          document.getElementById("importCancel").onclick = function () {
            modal.classList.add("hidden");
          };

          // Cerrar modal al hacer clic fuera
          window.onclick = function (event) {
            if (event.target === modal) {
              modal.classList.add("hidden");
            }
          };
          _context14.n = 3;
          break;
        case 2:
          _context14.p = 2;
          _t15 = _context14.v;
          console.error("Error al abrir modal de importación:", _t15);
          showNotification("Error al abrir el modal de importación", "error");
        case 3:
          return _context14.a(2);
      }
    }, _callee14, null, [[0, 2]]);
  }));
  return _importData.apply(this, arguments);
}
function executeCreateSheet() {
  return _executeCreateSheet.apply(this, arguments);
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
function _executeCreateSheet() {
  _executeCreateSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee15() {
    var importType, importError, _t16;
    return _regenerator().w(function (_context15) {
      while (1) switch (_context15.p = _context15.n) {
        case 0:
          _context15.p = 0;
          importType = document.getElementById("importType").value;
          importError = document.getElementById("importError"); // Validar que se haya seleccionado una opción
          if (importType) {
            _context15.n = 1;
            break;
          }
          importError.textContent = "Debe seleccionar un tipo de importación";
          importError.classList.remove("hidden");
          return _context15.a(2);
        case 1:
          // Ocultar mensaje de error si había alguno
          importError.classList.add("hidden");
          console.log("Creando hoja para tipo:", importType);

          // Cerrar el modal
          document.getElementById("importModal").classList.add("hidden");

          // Crear la hoja según el tipo
          if (!(importType === "movimientos")) {
            _context15.n = 4;
            break;
          }
          // Cargar todos los datos necesarios antes de crear la hoja
          showNotification("Descargando datos necesarios...", "info");
          _context15.n = 2;
          return Promise.all([loadAccounts(), loadFlowCodes(), loadBudgetCodes(), loadCurrencies(), loadQuotationPlaces()]);
        case 2:
          _context15.n = 3;
          return createMovimientosSheet();
        case 3:
          _context15.n = 5;
          break;
        case 4:
          if (importType === "flujos") {
            // TODO: Implementar creación de hoja para flujos
            showNotification("Funcionalidad de creaci\xF3n de hoja para flujos en desarrollo", "info");
          }
        case 5:
          _context15.n = 7;
          break;
        case 6:
          _context15.p = 6;
          _t16 = _context15.v;
          console.error("Error en executeCreateSheet:", _t16);
          showNotification("Error al crear la hoja", "error");
        case 7:
          return _context15.a(2);
      }
    }, _callee15, null, [[0, 6]]);
  }));
  return _executeCreateSheet.apply(this, arguments);
}
function executeImport() {
  return _executeImport.apply(this, arguments);
}
/**
 * Lee los datos de la hoja "Movimientos" en Excel
 * 
 * @async
 * @returns {Promise<Array<Object>>} Array de objetos con los datos de cada fila
 * @throws {Error} Si la hoja no existe o hay problemas al leer los datos
 */
function _executeImport() {
  _executeImport = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee16() {
    var importType, importError, _t17;
    return _regenerator().w(function (_context16) {
      while (1) switch (_context16.p = _context16.n) {
        case 0:
          _context16.p = 0;
          importType = document.getElementById("importType").value;
          importError = document.getElementById("importError"); // Validar que se haya seleccionado una opción
          if (importType) {
            _context16.n = 1;
            break;
          }
          importError.textContent = "Debe seleccionar un tipo de importación";
          importError.classList.remove("hidden");
          return _context16.a(2);
        case 1:
          // Ocultar mensaje de error si había alguno
          importError.classList.add("hidden");
          console.log("Importación de tipo:", importType);

          // Cerrar el modal
          document.getElementById("importModal").classList.add("hidden");

          // Ejecutar importación según el tipo
          if (!(importType === "movimientos")) {
            _context16.n = 3;
            break;
          }
          _context16.n = 2;
          return importMovimientosToOData();
        case 2:
          _context16.n = 4;
          break;
        case 3:
          if (importType === "flujos") {
            showNotification("Funcionalidad de importaci\xF3n de flujos en desarrollo", "info");
          }
        case 4:
          _context16.n = 6;
          break;
        case 5:
          _context16.p = 5;
          _t17 = _context16.v;
          console.error("Error en executeImport:", _t17);
          showNotification("Error al preparar la importación", "error");
        case 6:
          return _context16.a(2);
      }
    }, _callee16, null, [[0, 5]]);
  }));
  return _executeImport.apply(this, arguments);
}
function readMovimientosSheet() {
  return _readMovimientosSheet.apply(this, arguments);
}
/**
 * Valida un registro de movimiento según los criterios de la Historia de Usuario
 * 
 * @param {Object} record - Registro a validar
 * @param {number} rowNumber - Número de fila (para mensajes de error)
 * @returns {Object} Objeto con {isValid: boolean, errors: string[], errorFields: string[]}
 */
function _readMovimientosSheet() {
  _readMovimientosSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee18() {
    return _regenerator().w(function (_context18) {
      while (1) switch (_context18.n) {
        case 0:
          _context18.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref6 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee17(context) {
              var sheet, usedRange, values, headers, records, i, row, record, j, header, value, _t18;
              return _regenerator().w(function (_context17) {
                while (1) switch (_context17.p = _context17.n) {
                  case 0:
                    _context17.p = 0;
                    sheet = context.workbook.worksheets.getItem("Movimientos");
                    usedRange = sheet.getUsedRange();
                    usedRange.load(["values", "rowCount"]);
                    _context17.n = 1;
                    return context.sync();
                  case 1:
                    values = usedRange.values;
                    if (!(values.length <= 1)) {
                      _context17.n = 2;
                      break;
                    }
                    throw new Error("La hoja no contiene datos para importar");
                  case 2:
                    // La primera fila son las cabeceras
                    headers = values[0];
                    records = []; // Procesar cada fila de datos (empezando desde la fila 2)
                    for (i = 1; i < values.length; i++) {
                      row = values[i];
                      record = {}; // Mapear cada columna a su campo correspondiente
                      for (j = 0; j < headers.length; j++) {
                        header = headers[j];
                        value = row[j]; // Saltar valores vacíos para campos opcionales
                        if (value !== null && value !== undefined && value !== "") {
                          record[header] = value;
                        }
                      }

                      // Solo agregar registros que tengan al menos un campo
                      if (Object.keys(record).length > 0) {
                        records.push(record);
                      }
                    }
                    console.log("Le\xEDdos ".concat(records.length, " registros de la hoja Movimientos"));
                    return _context17.a(2, records);
                  case 3:
                    _context17.p = 3;
                    _t18 = _context17.v;
                    if (!_t18.message.includes("ItemNotFound")) {
                      _context17.n = 4;
                      break;
                    }
                    throw new Error("No existe la hoja 'Movimientos'. Debe crearla primero.");
                  case 4:
                    throw _t18;
                  case 5:
                    return _context17.a(2);
                }
              }, _callee17, null, [[0, 3]]);
            }));
            return function (_x4) {
              return _ref6.apply(this, arguments);
            };
          }());
        case 1:
          return _context18.a(2, _context18.v);
      }
    }, _callee18);
  }));
  return _readMovimientosSheet.apply(this, arguments);
}
function validateMovimientoRecord(record, rowNumber) {
  var errors = [];
  var errorFields = []; // Campos con error para marcar en rojo

  // Validar Status (requerido)
  if (!record.Status || record.Status.toString().trim() === "") {
    errors.push("Fila ".concat(rowNumber, ": El campo Status es obligatorio"));
    errorFields.push('Status');
  }

  // Validar IsDebit (requerido)
  if (record.IsDebit === null || record.IsDebit === undefined || record.IsDebit === "") {
    errors.push("Fila ".concat(rowNumber, ": El campo IsDebit es obligatorio"));
    errorFields.push('IsDebit');
  }

  // Validar Amount (distinto de 0)
  if (!record.Amount || parseFloat(record.Amount) === 0) {
    errors.push("Fila ".concat(rowNumber, ": El campo Amount debe ser distinto de 0"));
    errorFields.push('Amount');
  }

  // Validar ValueDate (requerido y formato válido)
  if (!record.ValueDate) {
    errors.push("Fila ".concat(rowNumber, ": El campo ValueDate es obligatorio"));
    errorFields.push('ValueDate');
  } else if (!isValidDate(record.ValueDate)) {
    errors.push("Fila ".concat(rowNumber, ": El campo ValueDate tiene formato inv\xE1lido. Use dd/mm/yyyy"));
    errorFields.push('ValueDate');
  }

  // Validar TrnAmount (distinto de 0)
  if (!record.TrnAmount || parseFloat(record.TrnAmount) === 0) {
    errors.push("Fila ".concat(rowNumber, ": El campo TrnAmount debe ser distinto de 0"));
    errorFields.push('TrnAmount');
  }

  // Validar TrnDate (requerido y formato válido)
  if (!record.TrnDate) {
    errors.push("Fila ".concat(rowNumber, ": El campo TrnDate es obligatorio"));
    errorFields.push('TrnDate');
  } else if (!isValidDate(record.TrnDate)) {
    errors.push("Fila ".concat(rowNumber, ": El campo TrnDate tiene formato inv\xE1lido. Use dd/mm/yyyy"));
    errorFields.push('TrnDate');
  }

  // Validar Number (mayor o igual a 1)
  if (!record.Number || parseInt(record.Number) < 1) {
    errors.push("Fila ".concat(rowNumber, ": El campo Number debe ser >= 1"));
    errorFields.push('Number');
  }

  // Validar Account (requerido)
  if (!record.Account || record.Account.toString().trim() === "") {
    errors.push("Fila ".concat(rowNumber, ": El campo Account es obligatorio"));
    errorFields.push('Account');
  }

  // Validar campos booleanos requeridos
  var booleanFields = ['UseInBalanceVal', 'UseInBalanceTrn', 'Interco', 'UseIntercoChart', 'IsManualFee'];
  booleanFields.forEach(function (field) {
    if (record[field] === null || record[field] === undefined || record[field] === "") {
      errors.push("Fila ".concat(rowNumber, ": El campo ").concat(field, " es obligatorio"));
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
function markErrorCells(_x2) {
  return _markErrorCells.apply(this, arguments);
}
/**
 * Valida si un valor es una fecha válida en formato dd/mm/yyyy o número serial de Excel
 * 
 * @param {any} value - Valor a validar
 * @returns {boolean} true si es una fecha válida
 */
function _markErrorCells() {
  _markErrorCells = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee20(validationResults) {
    return _regenerator().w(function (_context20) {
      while (1) switch (_context20.n) {
        case 0:
          _context20.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref7 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee19(context) {
              var sheet, headerRange, headers, _t19;
              return _regenerator().w(function (_context19) {
                while (1) switch (_context19.p = _context19.n) {
                  case 0:
                    _context19.p = 0;
                    sheet = context.workbook.worksheets.getItem("Movimientos"); // Obtener las cabeceras para saber qué columna es cada campo
                    headerRange = sheet.getRange("A1:S1");
                    headerRange.load("values");
                    _context19.n = 1;
                    return context.sync();
                  case 1:
                    headers = headerRange.values[0]; // Procesar cada resultado de validación que tenga errores
                    validationResults.forEach(function (result) {
                      if (!result.isValid && result.errorFields.length > 0) {
                        result.errorFields.forEach(function (fieldName) {
                          // Encontrar el índice de la columna para este campo
                          var columnIndex = headers.indexOf(fieldName);
                          if (columnIndex !== -1) {
                            // Convertir índice de columna a letra (A, B, C, etc.)
                            var columnLetter = String.fromCharCode(65 + columnIndex);
                            var cellAddress = "".concat(columnLetter).concat(result.rowNumber);

                            // Marcar solo el contorno de la celda en rojo
                            var errorCell = sheet.getRange(cellAddress);
                            ["EdgeTop", "EdgeBottom", "EdgeLeft", "EdgeRight"].forEach(function (edge) {
                              var border = errorCell.format.borders.getItem(edge);
                              border.style = "Continuous";
                              border.color = "#CC0000"; // Borde rojo
                            });
                          }
                        });
                      }
                    });
                    _context19.n = 2;
                    return context.sync();
                  case 2:
                    console.log("Celdas con errores marcadas con borde rojo");
                    _context19.n = 4;
                    break;
                  case 3:
                    _context19.p = 3;
                    _t19 = _context19.v;
                    console.error("Error al marcar celdas:", _t19);
                  case 4:
                    return _context19.a(2);
                }
              }, _callee19, null, [[0, 3]]);
            }));
            return function (_x5) {
              return _ref7.apply(this, arguments);
            };
          }());
        case 1:
          return _context20.a(2);
      }
    }, _callee20);
  }));
  return _markErrorCells.apply(this, arguments);
}
function isValidDate(value) {
  if (!value) return false;

  // Si es un número (serial de Excel), verificar que esté en rango válido
  if (typeof value === 'number') {
    return value > 0 && value < 2958466; // Rango válido de Excel (1900-9999)
  }

  // Si es string, validar formato dd/mm/yyyy
  if (typeof value === 'string') {
    var datePattern = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    var match = value.match(datePattern);
    if (!match) return false;
    var day = parseInt(match[1]);
    var month = parseInt(match[2]);
    var year = parseInt(match[3]);

    // Validar rangos
    if (month < 1 || month > 12) return false;
    if (day < 1 || day > 31) return false;
    if (year < 1900 || year > 9999) return false;

    // Validar días por mes
    var daysInMonth = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    var isLeapYear = year % 4 === 0 && year % 100 !== 0 || year % 400 === 0;
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
  var date;

  // Si es un número serial de Excel
  if (typeof dateValue === 'number') {
    // Excel cuenta los días desde 30/12/1899
    var excelEpoch = new Date(1899, 11, 30);
    var msPerDay = 24 * 60 * 60 * 1000;
    date = new Date(excelEpoch.getTime() + dateValue * msPerDay);
  }
  // Si es string en formato dd/mm/yyyy
  else if (typeof dateValue === 'string') {
    var parts = dateValue.split('/');
    var day = parseInt(parts[0]);
    var month = parseInt(parts[1]) - 1; // JavaScript months son 0-indexed
    var year = parseInt(parts[2]);
    date = new Date(year, month, day);
  } else {
    throw new Error("Formato de fecha no soportado: ".concat(dateValue));
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
  var payload = {};

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
  var accountCode = record.Account.toString().trim();
  var account = cachedAccounts.find(function (acc) {
    return acc.Code === accountCode;
  });
  if (account) {
    payload.Entity["Account@odata.bind"] = "Account2CashSet(".concat(account.Id, ")");
  } else {
    throw new Error("No se encontr\xF3 la cuenta con c\xF3digo: ".concat(accountCode));
  }

  // Mapear BudgetCode a ID (opcional)
  if (record.BudgetCode) {
    var budgetCode = record.BudgetCode.toString().trim();
    var budget = cachedBudgetCodes.find(function (bc) {
      return bc.Code === budgetCode;
    });
    if (budget) {
      payload.Entity["BudgetCode@odata.bind"] = "BudgetCodeSet(".concat(budget.Id, ")");
    } else {
      console.warn("No se encontr\xF3 el c\xF3digo presupuestario: ".concat(budgetCode));
    }
  }

  // Mapear FlowCode a ID (opcional)
  if (record.FlowCode) {
    var flowCode = record.FlowCode.toString().trim();
    var flow = cachedFlowCodes.find(function (fc) {
      return fc.Code === flowCode;
    });
    if (flow) {
      payload.Entity["FlowCode@odata.bind"] = "FlowCodeSet(".concat(flow.Id, ")");
    } else {
      console.warn("No se encontr\xF3 el flujo de caja: ".concat(flowCode));
    }
  }

  // Mapear TrnCurrency a ID (opcional)
  if (record.TrnCurrency) {
    var currencyCode = record.TrnCurrency.toString().trim();
    var currency = cachedCurrencies.find(function (c) {
      return c.Code === currencyCode;
    });
    if (currency) {
      payload.Entity["TrnCurrency@odata.bind"] = "CurrencySet('".concat(currency.Id, "')");
    } else {
      console.warn("No se encontr\xF3 la divisa: ".concat(currencyCode));
    }
  }

  // Mapear QuotationPlace a ID (opcional)
  if (record.QuotationPlace) {
    var quotationDesc = record.QuotationPlace.toString().trim();
    var quotation = cachedQuotationPlaces.find(function (qp) {
      return qp.Description === quotationDesc;
    });
    if (quotation) {
      payload.Entity["QuotationPlace@odata.bind"] = "QuotationPlaceSet(".concat(quotation.Id, ")");
    } else {
      console.warn("No se encontr\xF3 el lugar de cotizaci\xF3n: ".concat(quotationDesc));
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
  var requests = records.map(function (record, index) {
    // Construir el payload individual usando la función existente
    var payload = buildMovimientoJSON(record);
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
  return {
    requests: requests
  };
}

/**
 * Importa movimientos desde la hoja de Excel al servidor OData
 * 
 * @async
 */
function importMovimientosToOData() {
  return _importMovimientosToOData.apply(this, arguments);
}
/**
 * Crea una hoja de Excel para importar Movimientos con las cabeceras predefinidas
 * 
 * @async
 * @throws {Error} Si hay problemas al crear la hoja o escribir las cabeceras
 */
function _importMovimientosToOData() {
  _importMovimientosToOData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee21() {
    var records, validationResults, allErrors, invalidRecords, errorMessage, payload, isDevelopment, VERCEL_PROXY, endpoint, response, result, errorText, _errorMessage, batchPayload, _isDevelopment, _VERCEL_PROXY, _endpoint, _response, _result, responses, successCount, errorCount, _errorMessage2, _errorText, _errorMessage3, _t20;
    return _regenerator().w(function (_context21) {
      while (1) switch (_context21.p = _context21.n) {
        case 0:
          _context21.p = 0;
          showNotification("Preparando importación...", "info");

          // 1. Cargar todos los datos de referencia necesarios
          console.log("Cargando datos de referencia...");
          _context21.n = 1;
          return Promise.all([loadAccounts(), loadFlowCodes(), loadBudgetCodes(), loadCurrencies(), loadQuotationPlaces()]);
        case 1:
          // 2. Leer datos de la hoja
          console.log("Leyendo datos de la hoja Movimientos...");
          _context21.n = 2;
          return readMovimientosSheet();
        case 2:
          records = _context21.v;
          if (!(records.length === 0)) {
            _context21.n = 3;
            break;
          }
          showNotification("No hay datos para importar", "error");
          return _context21.a(2);
        case 3:
          // 3. Validar todos los registros
          console.log("Validando ".concat(records.length, " registros..."));
          validationResults = records.map(function (record, index) {
            return validateMovimientoRecord(record, index + 2);
          } // +2 porque la fila 1 son cabeceras
          );
          allErrors = validationResults.flatMap(function (result) {
            return result.errors;
          });
          if (!(allErrors.length > 0)) {
            _context21.n = 5;
            break;
          }
          console.error("Errores de validación:", allErrors);

          // Marcar celdas con errores en rojo
          _context21.n = 4;
          return markErrorCells(validationResults);
        case 4:
          // Crear mensaje detallado con los campos problemáticos
          invalidRecords = validationResults.filter(function (r) {
            return !r.isValid;
          });
          errorMessage = "\u26A0\uFE0F Validaci\xF3n fallida: ".concat(allErrors.length, " error(es) encontrado(s)\n\n");
          invalidRecords.forEach(function (result) {
            errorMessage += "\uD83D\uDCCD Fila ".concat(result.rowNumber, ":\n");
            errorMessage += "   Campos con problema: ".concat(result.errorFields.join(', '), "\n\n");
          });
          errorMessage += "Las celdas con errores han sido marcadas en rojo. Corríjalas e intente de nuevo.";

          // Mostrar mensaje en notificación
          showNotification("Errores de validación encontrados", "error");

          // Mostrar mensaje detallado en panel modal
          showErrorDetails(errorMessage);
          allErrors.forEach(function (error) {
            return console.error(error);
          });
          return _context21.a(2);
        case 5:
          if (!(records.length === 1)) {
            _context21.n = 11;
            break;
          }
          console.log("Enviando único registro al servidor...");
          payload = buildMovimientoJSON(records[0]);
          console.log("Payload JSON:", JSON.stringify(payload, null, 2));
          isDevelopment = window.location.hostname === 'localhost';
          VERCEL_PROXY = isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
          endpoint = "".concat(VERCEL_PROXY, "CashFlowDtoWithExtendersSet");
          console.log("Enviando POST a:", endpoint);
          _context21.n = 6;
          return authenticatedFetch(endpoint, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json; charset=utf-8',
              'Accept': 'application/json; charset=utf-8'
            },
            body: JSON.stringify(payload)
          });
        case 6:
          response = _context21.v;
          if (!response.ok) {
            _context21.n = 8;
            break;
          }
          _context21.n = 7;
          return response.json();
        case 7:
          result = _context21.v;
          console.log("Respuesta del servidor:", result);
          showNotification("✅ Movimiento importado exitosamente al servidor OData", "success");
          _context21.n = 10;
          break;
        case 8:
          _context21.n = 9;
          return response.text();
        case 9:
          errorText = _context21.v;
          console.error("Error del servidor:", response.status, errorText);

          // Crear mensaje de error detallado
          _errorMessage = "\u274C Error al a\xF1adir el movimiento\n\n";
          _errorMessage += "C\xF3digo de error: ".concat(response.status, "\n");
          _errorMessage += "Detalles: ".concat(errorText.substring(0, 200), "\n\n");
          _errorMessage += "\uD83D\uDCA1 Sugerencias:\n";
          _errorMessage += "- Verifique que el campo Status sea \"Actual\" (no \"Active\")\n";
          _errorMessage += "- Revise que todos los c\xF3digos de cuenta, flujo y presupuesto sean v\xE1lidos\n";
          _errorMessage += "- Compruebe que las fechas est\xE9n en formato correcto\n";
          showNotification("Error al importar movimiento", "error");
          showErrorDetails(_errorMessage);
        case 10:
          _context21.n = 16;
          break;
        case 11:
          // Múltiples registros: usar OData $batch
          console.log("Enviando ".concat(records.length, " registros en lote al servidor..."));
          batchPayload = buildBatchRequestJSON(records);
          console.log("Batch Payload JSON:", JSON.stringify(batchPayload, null, 2));
          _isDevelopment = window.location.hostname === 'localhost';
          _VERCEL_PROXY = _isDevelopment ? '/odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
          _endpoint = "".concat(_VERCEL_PROXY, "$batch");
          console.log("Enviando POST batch a:", _endpoint);
          _context21.n = 12;
          return authenticatedFetch(_endpoint, {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json; charset=utf-8',
              'Accept': 'application/json; charset=utf-8'
            },
            body: JSON.stringify(batchPayload)
          });
        case 12:
          _response = _context21.v;
          if (!_response.ok) {
            _context21.n = 14;
            break;
          }
          _context21.n = 13;
          return _response.json();
        case 13:
          _result = _context21.v;
          console.log("Respuesta del servidor (batch):", _result);

          // Analizar respuesta batch para ver cuántos tuvieron éxito
          responses = _result.responses || [];
          successCount = responses.filter(function (r) {
            return r.status >= 200 && r.status < 300;
          }).length;
          errorCount = responses.length - successCount;
          if (errorCount === 0) {
            showNotification("\u2705 ".concat(successCount, " movimientos importados exitosamente"), "success");
          } else {
            _errorMessage2 = "\u26A0\uFE0F Importaci\xF3n parcial:\n\n";
            _errorMessage2 += "\u2705 Exitosos: ".concat(successCount, "\n");
            _errorMessage2 += "\u274C Fallidos: ".concat(errorCount, "\n\n");
            _errorMessage2 += "Detalles de errores:\n";
            responses.forEach(function (resp, idx) {
              if (resp.status >= 300) {
                _errorMessage2 += "\n\u2022 Registro ".concat(idx + 1, " (fila ").concat(idx + 2, "): Error ").concat(resp.status, "\n");
                if (resp.body && resp.body.error) {
                  _errorMessage2 += "  ".concat(resp.body.error.message, "\n");
                }
              }
            });
            showNotification("Importaci\xF3n completada con errores", "warning");
            showErrorDetails(_errorMessage2);
          }
          _context21.n = 16;
          break;
        case 14:
          _context21.n = 15;
          return _response.text();
        case 15:
          _errorText = _context21.v;
          console.error("Error del servidor (batch):", _response.status, _errorText);
          _errorMessage3 = "\u274C Error al enviar el lote de movimientos\n\n";
          _errorMessage3 += "C\xF3digo de error: ".concat(_response.status, "\n");
          _errorMessage3 += "Detalles: ".concat(_errorText.substring(0, 300), "\n\n");
          _errorMessage3 += "\uD83D\uDCA1 Sugerencias:\n";
          _errorMessage3 += "- Verifique que todos los registros tengan datos v\xE1lidos\n";
          _errorMessage3 += "- Compruebe la conectividad con el servidor OData\n";
          _errorMessage3 += "- Revise los logs de consola para m\xE1s detalles\n";
          showNotification("Error al importar lote de movimientos", "error");
          showErrorDetails(_errorMessage3);
        case 16:
          _context21.n = 18;
          break;
        case 17:
          _context21.p = 17;
          _t20 = _context21.v;
          console.error("Error durante la importación:", _t20);
          showNotification("Error durante la importación: " + _t20.message, "error");
        case 18:
          return _context21.a(2);
      }
    }, _callee21, null, [[0, 17]]);
  }));
  return _importMovimientosToOData.apply(this, arguments);
}
function createMovimientosSheet() {
  return _createMovimientosSheet.apply(this, arguments);
}
function _createMovimientosSheet() {
  _createMovimientosSheet = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee23() {
    var _t22;
    return _regenerator().w(function (_context23) {
      while (1) switch (_context23.p = _context23.n) {
        case 0:
          _context23.p = 0;
          _context23.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref8 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee22(context) {
              var application, sheetName, existingSheet, sheet, headers, headerRange, valueDateColumn, valueDateValidation, trnDateColumn, trnDateValidation, accountColumn, accountValidation, accountCodes, budgetCodeColumn, budgetCodeValidation, budgetCodes, flowCodeColumn, flowCodeValidation, flowCodes, currencyColumn, currencyValidation, currencyCodes, quotationPlaceColumn, quotationPlaceValidation, quotationDescriptions, statusColumn, statusValidation, booleanColumns, _t21;
              return _regenerator().w(function (_context22) {
                while (1) switch (_context22.p = _context22.n) {
                  case 0:
                    application = context.workbook.application;
                    application.suspendScreenUpdatingUntilNextSync();
                    sheetName = "Movimientos"; // Verificar si la hoja existe y eliminarla
                    _context22.p = 1;
                    existingSheet = context.workbook.worksheets.getItem(sheetName);
                    existingSheet.delete();
                    _context22.n = 2;
                    return context.sync();
                  case 2:
                    console.log("Hoja existente '".concat(sheetName, "' eliminada"));
                    _context22.n = 4;
                    break;
                  case 3:
                    _context22.p = 3;
                    _t21 = _context22.v;
                    console.log("La hoja '".concat(sheetName, "' no existe, se crear\xE1 una nueva"));
                  case 4:
                    // Crear la hoja
                    sheet = context.workbook.worksheets.add(sheetName);
                    sheet.load(["protection", "name"]);
                    _context22.n = 5;
                    return context.sync();
                  case 5:
                    if (!sheet.protection.protected) {
                      _context22.n = 6;
                      break;
                    }
                    throw new Error("La hoja está protegida. No se pueden escribir datos.");
                  case 6:
                    console.log("Hoja creada: ".concat(sheetName));

                    // Definir las cabeceras basadas en el JSON
                    headers = ["TERCERO", "Status", "IsDebit", "Amount", "ValueDate", "TrnAmount", "TrnDate", "Number", "Description", "Account", "BudgetCode", "FlowCode", "TrnCurrency", "QuotationPlace", "UseInBalanceVal", "UseInBalanceTrn", "Interco", "UseIntercoChart", "IsManualFee"]; // Escribir las cabeceras en la primera fila
                    headerRange = sheet.getRange("A1:".concat(String.fromCharCode(64 + headers.length), "1"));
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
                    valueDateColumn = sheet.getRange("E2:E1048576");
                    valueDateValidation = valueDateColumn.dataValidation;
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
                    trnDateColumn = sheet.getRange("G2:G1048576");
                    trnDateValidation = trnDateColumn.dataValidation;
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
                      accountColumn = sheet.getRange("J2:J1048576");
                      accountValidation = accountColumn.dataValidation; // Crear la lista de valores separados por comas (solo los códigos de cuenta)
                      accountCodes = cachedAccounts.map(function (acc) {
                        return acc.Code;
                      }).join(",");
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
                      console.log("Dropdown de cuentas configurado con ".concat(cachedAccounts.length, " cuentas"));
                    } else {
                      console.warn("No hay cuentas en caché. Descargue las cuentas primero para habilitar el dropdown.");
                    }

                    // Agregar dropdown para la columna BudgetCode (columna K, índice 10)
                    if (cachedBudgetCodes && cachedBudgetCodes.length > 0) {
                      budgetCodeColumn = sheet.getRange("K2:K1048576");
                      budgetCodeValidation = budgetCodeColumn.dataValidation;
                      budgetCodes = cachedBudgetCodes.map(function (bc) {
                        return bc.Code;
                      }).join(",");
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
                      console.log("Dropdown de c\xF3digos presupuestarios configurado con ".concat(cachedBudgetCodes.length, " c\xF3digos"));
                    }

                    // Agregar dropdown para la columna FlowCode (columna L, índice 11)
                    if (cachedFlowCodes && cachedFlowCodes.length > 0) {
                      flowCodeColumn = sheet.getRange("L2:L1048576");
                      flowCodeValidation = flowCodeColumn.dataValidation;
                      flowCodes = cachedFlowCodes.map(function (fc) {
                        return fc.Code;
                      }).join(",");
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
                      console.log("Dropdown de flujos de caja configurado con ".concat(cachedFlowCodes.length, " flujos"));
                    }

                    // Agregar dropdown para la columna TrnCurrency (columna M, índice 12)
                    if (cachedCurrencies && cachedCurrencies.length > 0) {
                      currencyColumn = sheet.getRange("M2:M1048576");
                      currencyValidation = currencyColumn.dataValidation; // Crear la lista de códigos de divisa separados por comas
                      currencyCodes = cachedCurrencies.map(function (c) {
                        return c.Code;
                      }).join(",");
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
                      console.log("Dropdown de divisas configurado con ".concat(cachedCurrencies.length, " divisas"));
                    } else {
                      console.warn("No hay divisas en caché. Descargue las divisas primero para habilitar el dropdown.");
                    }

                    // Agregar dropdown para la columna QuotationPlace (columna N, índice 13)
                    if (cachedQuotationPlaces && cachedQuotationPlaces.length > 0) {
                      quotationPlaceColumn = sheet.getRange("N2:N1048576");
                      quotationPlaceValidation = quotationPlaceColumn.dataValidation; // Crear la lista de descripciones de lugares de cotización separadas por comas
                      quotationDescriptions = cachedQuotationPlaces.map(function (qp) {
                        return qp.Description;
                      }).join(",");
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
                      console.log("Dropdown de lugares de cotizaci\xF3n configurado con ".concat(cachedQuotationPlaces.length, " lugares"));
                    } else {
                      console.warn("No hay lugares de cotización en caché. Descargue primero para habilitar el dropdown.");
                    }

                    // Agregar dropdown para Status (columna B, índice 1)
                    statusColumn = sheet.getRange("B2:B1048576");
                    statusValidation = statusColumn.dataValidation;
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
                    booleanColumns = [{
                      range: "C2:C1048576",
                      name: "IsDebit",
                      title: "¿Es débito?"
                    }, {
                      range: "O2:O1048576",
                      name: "UseInBalanceVal",
                      title: "¿Usar en balance de valor?"
                    }, {
                      range: "P2:P1048576",
                      name: "UseInBalanceTrn",
                      title: "¿Usar en balance de transacción?"
                    }, {
                      range: "Q2:Q1048576",
                      name: "Interco",
                      title: "¿Es intercompañía?"
                    }, {
                      range: "R2:R1048576",
                      name: "UseIntercoChart",
                      title: "¿Usar plan intercompañía?"
                    }, {
                      range: "S2:S1048576",
                      name: "IsManualFee",
                      title: "¿Es comisión manual?"
                    }];
                    booleanColumns.forEach(function (col) {
                      var column = sheet.getRange(col.range);
                      var validation = column.dataValidation;
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
                    _context22.n = 7;
                    return context.sync();
                  case 7:
                    showNotification("Hoja 'Movimientos' creada con controles de validación", "success");
                  case 8:
                    return _context22.a(2);
                }
              }, _callee22, null, [[1, 3]]);
            }));
            return function (_x6) {
              return _ref8.apply(this, arguments);
            };
          }());
        case 1:
          _context23.n = 3;
          break;
        case 2:
          _context23.p = 2;
          _t22 = _context23.v;
          console.error("Error al crear hoja de Movimientos:", _t22);
          showNotification("Error al crear la hoja de Movimientos: " + _t22.message, "error");
        case 3:
          return _context23.a(2);
      }
    }, _callee23, null, [[0, 2]]);
  }));
  return _createMovimientosSheet.apply(this, arguments);
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ../../assets/logo-filled.png */ "./assets/logo-filled.png"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html lang=\"es\">\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Add-in Azurriga</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\r\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\r\n        <img width=\"90\" height=\"90\" src=\"" + ___HTML_LOADER_IMPORT_1___ + "\" alt=\"Azurriga\" title=\"Azurriga\" />\r\n         <!-- <h1 class=\"ms-font-su\">Welcome</h1>-->\r\n    </header>\r\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\r\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" rel=\"noopener noreferrer\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\r\n    </section>\r\n    <main id=\"app-body\" class=\"ms-welcome__main hidden\">\r\n        <div class=\"ms-welcome__buttons\">\r\n            <div role=\"button\" id=\"login\" class=\"ms-Button ms-Button--primary\">\r\n                <span class=\"ms-Button-label\">Login</span>\r\n            </div>\r\n            <div role=\"button\" id=\"download\" class=\"ms-Button ms-Button--primary is-disabled\" disabled>\r\n                <span class=\"ms-Button-label\">Download</span>\r\n            </div>\r\n            <div role=\"button\" id=\"import\" class=\"ms-Button ms-Button--primary is-disabled\" disabled>\r\n                <span class=\"ms-Button-label\">Import</span>\r\n            </div>\r\n        </div>\r\n        <p><label id=\"item-subject\"></label></p>\r\n    </main>\r\n\r\n    <!-- Popup de Notificación -->\r\n    <div id=\"notificationPopup\" class=\"notification-popup hidden\">\r\n        <div class=\"notification-content\">\r\n            <span id=\"notificationMessage\"></span>\r\n        </div>\r\n    </div>\r\n\r\n    <!-- Área de Errores Detallados -->\r\n    <div id=\"errorDetailsPanel\" class=\"error-details-panel hidden\">\r\n        <div class=\"error-details-content\">\r\n            <div class=\"error-details-header\">\r\n                <h3>Errores de Validación</h3>\r\n                <button id=\"closeErrorDetails\" class=\"close-button\">&times;</button>\r\n            </div>\r\n            <div id=\"errorDetailsMessage\" class=\"error-details-message\"></div>\r\n        </div>\r\n    </div>\r\n\r\n    <!-- Modal de Login -->\r\n    <div id=\"loginModal\" class=\"modal hidden\">\r\n        <div class=\"modal-content ms-depth-4\">\r\n            <h2 class=\"ms-font-xl\">Iniciar Sesión</h2>\r\n            <div class=\"form-group\">\r\n                <label class=\"ms-Label\">Usuario:</label>\r\n                <input type=\"text\" id=\"username\" class=\"ms-TextField-field\" placeholder=\"Ingrese su usuario\">\r\n            </div>\r\n            <div class=\"form-group\">\r\n                <label class=\"ms-Label\">Contraseña:</label>\r\n                <input type=\"password\" id=\"password\" class=\"ms-TextField-field\" placeholder=\"Ingrese su contraseña\">\r\n            </div>\r\n            <div id=\"loginError\" class=\"error-message hidden\"></div>\r\n            \r\n            <!-- Loading spinner -->\r\n            <div id=\"loginLoading\" class=\"loading-container hidden\">\r\n                <div class=\"spinner\"></div>\r\n                <p class=\"ms-font-m\">Conectando al servidor...</p>\r\n            </div>\r\n            \r\n            <div class=\"modal-buttons\">\r\n                <button class=\"ms-Button ms-Button--primary\" id=\"loginSubmit\">\r\n                    <span class=\"ms-Button-label\">Iniciar Sesión</span>\r\n                </button>\r\n                <button class=\"ms-Button\" id=\"loginCancel\">\r\n                    <span class=\"ms-Button-label\">Cancelar</span>\r\n                </button>\r\n            </div>\r\n        </div>\r\n    </div>\r\n\r\n    <!-- Modal de Descarga -->\r\n    <div id=\"downloadModal\" class=\"modal hidden\">\r\n        <div class=\"modal-content ms-depth-4\">\r\n            <h2 class=\"ms-font-xl\">Opciones de Descarga</h2>\r\n            \r\n            <!-- Selector de tipo de descarga -->\r\n            <div class=\"form-group\">\r\n                <label class=\"ms-Label\">Tipo de descarga:</label>\r\n                <select id=\"downloadType\" class=\"ms-Dropdown\" title=\"Seleccionar tipo de descarga\">\r\n                    <option value=\"cuentas\">Cuentas</option>\r\n                    <option value=\"flujos\">Flujos</option>\r\n                    <option value=\"codigos-presupuestarios\">Códigos Presupuestarios</option>\r\n                    <option value=\"divisas\">Divisas</option>\r\n                    <option value=\"cotizacion\">Cotización</option>\r\n                    <option value=\"movimientos\">Movimientos</option>\r\n                </select>\r\n            </div>\r\n\r\n            <!-- Selector de cantidad de registros -->\r\n            <div class=\"form-group\">\r\n                <label class=\"ms-Label\">Cantidad de registros:</label>\r\n                <select id=\"recordLimit\" class=\"ms-Dropdown\" title=\"Seleccionar cantidad de registros\">\r\n                    <option value=\"50\">50</option>\r\n                    <option value=\"75\">75</option>\r\n                    <option value=\"100\">100</option>\r\n                    <option value=\"500\">500</option>\r\n                    <option value=\"1000\">1000</option>\r\n                    <option value=\"all\">Todas</option>\r\n                </select>\r\n            </div>\r\n\r\n            <!-- Opciones específicas para Movimientos -->\r\n            <div id=\"movimientosOptions\" class=\"form-group hidden\">\r\n                <label class=\"ms-Label\">Seleccionar Cuenta:</label>\r\n                <select id=\"accountSelect\" class=\"ms-Dropdown\" title=\"Seleccionar cuenta\">\r\n                    <option value=\"\">Todas las cuentas</option>\r\n                </select>\r\n                \r\n                <label class=\"ms-Label\" style=\"margin-top: 15px;\">Filtrar por Fecha (ValueDate):</label>\r\n                <div style=\"display: flex; gap: 10px; margin-bottom: 15px;\">\r\n                    <div style=\"flex: 1;\">\r\n                        <label class=\"ms-Label\" style=\"font-size: 12px;\">Desde:</label>\r\n                        <input type=\"date\" id=\"dateFrom\" class=\"ms-TextField-field\" title=\"Fecha inicio\">\r\n                    </div>\r\n                    <div style=\"flex: 1;\">\r\n                        <label class=\"ms-Label\" style=\"font-size: 12px;\">Hasta:</label>\r\n                        <input type=\"date\" id=\"dateTo\" class=\"ms-TextField-field\" title=\"Fecha fin\">\r\n                    </div>\r\n                </div>\r\n                \r\n                <label class=\"ms-Label\" style=\"margin-top: 15px;\">Campos a incluir:</label>\r\n                <div class=\"checkbox-group\">\r\n                    <div class=\"ms-CheckBox\">\r\n                        <input type=\"checkbox\" id=\"fieldStatus\" value=\"Status\" checked>\r\n                        <label for=\"fieldStatus\">Status</label>\r\n                    </div>\r\n                    <div class=\"ms-CheckBox\">\r\n                        <input type=\"checkbox\" id=\"fieldIsDebit\" value=\"IsDebit\" checked>\r\n                        <label for=\"fieldIsDebit\">IsDebit</label>\r\n                    </div>\r\n                    <div class=\"ms-CheckBox\">\r\n                        <input type=\"checkbox\" id=\"fieldAmount\" value=\"Amount\" checked>\r\n                        <label for=\"fieldAmount\">Amount</label>\r\n                    </div>\r\n                    <div class=\"ms-CheckBox\">\r\n                        <input type=\"checkbox\" id=\"fieldValueDate\" value=\"ValueDate\" checked>\r\n                        <label for=\"fieldValueDate\">ValueDate</label>\r\n                    </div>\r\n                    <div class=\"ms-CheckBox\">\r\n                        <input type=\"checkbox\" id=\"fieldTrnDate\" value=\"TrnDate\" checked>\r\n                        <label for=\"fieldTrnDate\">TrnDate</label>\r\n                    </div>\r\n                    <div class=\"ms-CheckBox\">\r\n                        <input type=\"checkbox\" id=\"fieldDescription\" value=\"Description\" checked>\r\n                        <label for=\"fieldDescription\">Description</label>\r\n                    </div>\r\n                </div>\r\n            </div>\r\n\r\n            <div id=\"downloadError\" class=\"error-message hidden\"></div>\r\n            \r\n            <div class=\"modal-buttons\">\r\n                <button class=\"ms-Button ms-Button--primary\" id=\"downloadSubmit\">\r\n                    <span class=\"ms-Button-label\">Descargar</span>\r\n                </button>\r\n                <button class=\"ms-Button\" id=\"downloadCancel\">\r\n                    <span class=\"ms-Button-label\">Cancelar</span>\r\n                </button>\r\n            </div>\r\n        </div>\r\n    </div>\r\n\r\n    <!-- Modal de Importación -->\r\n    <div id=\"importModal\" class=\"modal hidden\">\r\n        <div class=\"modal-content ms-depth-4\">\r\n            <h2 class=\"ms-font-xl\">Opciones de Importación</h2>\r\n            \r\n            <p class=\"ms-font-m\" style=\"margin-bottom: 20px; color: #605e5c;\">\r\n                Seleccione el tipo de datos que desea importar desde Excel al servidor:\r\n            </p>\r\n            \r\n            <!-- Selector de tipo de importación -->\r\n            <div class=\"form-group\">\r\n                <label class=\"ms-Label\">Tipo de importación:</label>\r\n                <select id=\"importType\" class=\"ms-Dropdown\" title=\"Seleccionar tipo de importación\">\r\n                    <option value=\"\">-- Seleccione una opción --</option>\r\n                    <option value=\"flujos\">Flujos</option>\r\n                    <option value=\"movimientos\">Movimientos</option>\r\n                </select>\r\n            </div>\r\n\r\n            <div id=\"importError\" class=\"error-message hidden\"></div>\r\n            \r\n            <div class=\"modal-buttons\">\r\n                <button class=\"ms-Button ms-Button--primary\" id=\"importCreateSheet\">\r\n                    <span class=\"ms-Button-label\">Crear hoja</span>\r\n                </button>\r\n                <button class=\"ms-Button ms-Button--primary\" id=\"importSubmit\">\r\n                    <span class=\"ms-Button-label\">Importar</span>\r\n                </button>\r\n                <button class=\"ms-Button\" id=\"importCancel\">\r\n                    <span class=\"ms-Button-label\">Cancelar</span>\r\n                </button>\r\n            </div>\r\n        </div>\r\n    </div>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map