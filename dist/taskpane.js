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

module.exports = __webpack_require__.p + "66d7f36bafce4bc2ecfe.css";

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
      } else {
        movimientosOptions.classList.add("hidden");
      }
    };
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

                  // DESARROLLO: Usar proxy local (http://localhost:3002)
                  // PRODUCCIÓN: Usar proxy Vercel (https://excel-addin-azurriga.vercel.app)
                  isDevelopment = window.location.hostname === 'localhost';
                  baseUrl = isDevelopment ? 'http://localhost:3002/api/proxy?path=odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
                  console.log("Usando proxy ".concat(isDevelopment ? 'LOCAL' : 'VERCEL', ": ").concat(baseUrl));
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
function authenticatedFetch(_x) {
  return _authenticatedFetch.apply(this, arguments);
}
/**
 * Muestra el modal de configuración de descarga
 * 
 * Permite al usuario configurar:
 * - Tipo de descarga: Cuentas, Flujos de caja o Movimientos
 * - Límite de registros: 50, 100, 500 o todos
 * - Campos específicos (solo para Movimientos)
 * 
 * @async
 * @throws {Error} Si el usuario no está autenticado o hay problemas con el DOM
 */
function _authenticatedFetch() {
  _authenticatedFetch = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3(url) {
    var options,
      defaultOptions,
      _args3 = arguments;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          options = _args3.length > 1 && _args3[1] !== undefined ? _args3[1] : {};
          if (userCredentials.isLoggedIn) {
            _context3.n = 1;
            break;
          }
          throw new Error("Debe iniciar sesión primero");
        case 1:
          defaultOptions = {
            headers: {
              'Content-Type': 'application/json; charset=utf-8',
              'Accept': 'application/json; charset=utf-8',
              'Authorization': "Basic ".concat(btoa(userCredentials.username + ':' + userCredentials.password))
            }
          };
          return _context3.a(2, fetch(url, _objectSpread(_objectSpread({}, defaultOptions), options)));
      }
    }, _callee3);
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
  _showDownloadModal = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5() {
    var modal, _t4;
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.p = _context5.n) {
        case 0:
          _context5.p = 0;
          if (userCredentials.isLoggedIn) {
            _context5.n = 1;
            break;
          }
          showNotification("Debe iniciar sesión primero", "error");
          return _context5.a(2);
        case 1:
          modal = document.getElementById("downloadModal");
          modal.classList.remove("hidden");
          modal.style.display = "block";

          // Configurar botón de submit
          document.getElementById("downloadSubmit").onclick = /*#__PURE__*/_asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
            return _regenerator().w(function (_context4) {
              while (1) switch (_context4.n) {
                case 0:
                  _context4.n = 1;
                  return executeDownload();
                case 1:
                  return _context4.a(2);
              }
            }, _callee4);
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
          _context5.n = 3;
          break;
        case 2:
          _context5.p = 2;
          _t4 = _context5.v;
          console.error("Error al abrir modal de descarga:", _t4);
          showNotification("Error al abrir el modal de descarga", "error");
        case 3:
          return _context5.a(2);
      }
    }, _callee5, null, [[0, 2]]);
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
 * @param {string} [downloadType='cuentas'] - Tipo de datos: 'cuentas', 'flujos' o 'movimientos'
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
  _executeDownload = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee6() {
    var downloadType, recordLimit, selectedFields, checkboxes, _t5;
    return _regenerator().w(function (_context6) {
      while (1) switch (_context6.p = _context6.n) {
        case 0:
          _context6.p = 0;
          downloadType = document.getElementById("downloadType").value;
          recordLimit = document.getElementById("recordLimit").value; // Recoger campos seleccionados para Movimientos
          selectedFields = [];
          if (!(downloadType === "movimientos")) {
            _context6.n = 1;
            break;
          }
          checkboxes = document.querySelectorAll('#movimientosOptions input[type="checkbox"]:checked');
          selectedFields = Array.from(checkboxes).map(function (cb) {
            return cb.value;
          });
          if (!(selectedFields.length === 0)) {
            _context6.n = 1;
            break;
          }
          showNotification("Debe seleccionar al menos un campo", "error");
          return _context6.a(2);
        case 1:
          console.log("Tipo de descarga:", downloadType);
          console.log("Límite de registros:", recordLimit);
          console.log("Campos seleccionados:", selectedFields);

          // Cerrar el modal
          document.getElementById("downloadModal").classList.add("hidden");

          // Llamar a la función de descarga con los parámetros
          _context6.n = 2;
          return download(downloadType, recordLimit, selectedFields);
        case 2:
          _context6.n = 4;
          break;
        case 3:
          _context6.p = 3;
          _t5 = _context6.v;
          console.error("Error en executeDownload:", _t5);
          showNotification("Error al preparar la descarga", "error");
        case 4:
          return _context6.a(2);
      }
    }, _callee6, null, [[0, 3]]);
  }));
  return _executeDownload.apply(this, arguments);
}
function download() {
  return _download.apply(this, arguments);
}

/**
 * 📤 FUNCIÓN DE EJEMPLO - Importa datos desde Excel a un servidor externo
 * 
 * NOTA: Esta función es solo un ejemplo educativo. Usa un servidor de prueba
 * (jsonplaceholder.typicode.com) y no se utiliza en producción.
 * 
 * Flujo:
 * 1. Lee datos del rango A1:B2 de la hoja activa
 * 2. Valida que existan encabezados y datos
 * 3. Envía los datos por POST al servidor de prueba
 * 4. Crea una hoja "Resultado" con la respuesta del servidor
 * 5. Muestra notificación de éxito
 * 
 * Validaciones:
 * - Verifica que haya al menos 2 filas (encabezados + datos)
 * - Valida que los encabezados no estén vacíos
 * - Valida que haya al menos un dato
 * 
 * Características:
 * - Sistema de reintentos (3 intentos)
 * - Nombres de hoja únicos (Resultado, Resultado_1, Resultado_2, etc.)
 * - Formato visual para la hoja de resultados
 * 
 * @async
 * @throws {Error} Si no hay suficientes datos, faltan encabezados o hay problemas de conexión
 */
function _download() {
  _download = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee8() {
    var downloadType,
      recordLimit,
      selectedFields,
      errorMessage,
      _args8 = arguments,
      _t1;
    return _regenerator().w(function (_context8) {
      while (1) switch (_context8.p = _context8.n) {
        case 0:
          downloadType = _args8.length > 0 && _args8[0] !== undefined ? _args8[0] : 'cuentas';
          recordLimit = _args8.length > 1 && _args8[1] !== undefined ? _args8[1] : '50';
          selectedFields = _args8.length > 2 && _args8[2] !== undefined ? _args8[2] : [];
          _context8.p = 1;
          _context8.n = 2;
          return Excel.run(/*#__PURE__*/function () {
            var _ref3 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee7(context) {
              var application, isDevelopment, VERCEL_PROXY, endpoint, params, expandParam, separator, response, retries, data, records, sheetName, existingSheet, sheet, sheet1, formatDate, formatValue, headers, values, allFields, numRows, numCols, getColumnLetter, endColumn, range, headerRange, dateFields, idColIndex, idColLetter, idRange, _t6, _t7, _t8, _t9, _t0;
              return _regenerator().w(function (_context7) {
                while (1) switch (_context7.p = _context7.n) {
                  case 0:
                    application = context.workbook.application;
                    application.suspendScreenUpdatingUntilNextSync();

                    // DESARROLLO: Usar proxy local (http://localhost:3002)
                    // PRODUCCIÓN: Usar proxy Vercel (https://excel-addin-azurriga.vercel.app)
                    isDevelopment = window.location.hostname === 'localhost';
                    VERCEL_PROXY = isDevelopment ? 'http://localhost:3002/api/proxy?path=odata/' : 'https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/';
                    console.log("Download usando proxy ".concat(isDevelopment ? 'LOCAL' : 'VERCEL'));
                    endpoint = '';
                    _t6 = downloadType;
                    _context7.n = _t6 === 'cuentas' ? 1 : _t6 === 'flujos' ? 2 : _t6 === 'movimientos' ? 3 : 4;
                    break;
                  case 1:
                    endpoint = "".concat(VERCEL_PROXY, "AccountSet");
                    return _context7.a(3, 4);
                  case 2:
                    endpoint = "".concat(VERCEL_PROXY, "FlowCodeSet");
                    return _context7.a(3, 4);
                  case 3:
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

                    // Agregar $filter solo con Status
                    params.push("$filter=Status eq 'Actual'");

                    // Unir todos los parámetros (siempre codificados para proxy Vercel)
                    if (params.length > 0) {
                      endpoint += '%3F' + params.join('%26'); // %3F = ?, %26 = &
                    }
                    return _context7.a(3, 4);
                  case 4:
                    // Agregar límite de registros para Cuentas y Flujos
                    if (downloadType !== 'movimientos' && recordLimit !== 'all') {
                      separator = endpoint.includes('%3F') ? '%26' : '%3F';
                      endpoint += separator + "$top=".concat(recordLimit);
                    }
                    console.log("Descargando desde:", endpoint);
                    console.log("Usuario autenticado:", userCredentials.username);

                    // Intentar obtener los datos con autenticación
                    retries = 3;
                  case 5:
                    if (!(retries > 0)) {
                      _context7.n = 12;
                      break;
                    }
                    _context7.p = 6;
                    _context7.n = 7;
                    return authenticatedFetch(endpoint);
                  case 7:
                    response = _context7.v;
                    console.log("Respuesta recibida. Status:", response.status);
                    if (!response.ok) {
                      _context7.n = 8;
                      break;
                    }
                    return _context7.a(3, 12);
                  case 8:
                    _context7.n = 11;
                    break;
                  case 9:
                    _context7.p = 9;
                    _t7 = _context7.v;
                    console.error("Error en intento de fetch:", _t7);
                    retries--;
                    if (!(retries === 0)) {
                      _context7.n = 10;
                      break;
                    }
                    throw new Error('Error al obtener datos después de 3 intentos');
                  case 10:
                    _context7.n = 11;
                    return new Promise(function (resolve) {
                      return setTimeout(resolve, 1000);
                    });
                  case 11:
                    _context7.n = 5;
                    break;
                  case 12:
                    _context7.n = 13;
                    return response.json();
                  case 13:
                    data = _context7.v;
                    console.log("Datos recibidos:", data);

                    // Verificar que tengamos datos
                    if (!(!data || !data.value || data.value.length === 0)) {
                      _context7.n = 14;
                      break;
                    }
                    throw new Error("No se recibieron datos del servidor");
                  case 14:
                    records = data.value; // OData devuelve los datos en data.value
                    // Determinar el nombre de la hoja según el tipo de descarga
                    sheetName = '';
                    _t8 = downloadType;
                    _context7.n = _t8 === 'cuentas' ? 15 : _t8 === 'flujos' ? 16 : _t8 === 'movimientos' ? 17 : 18;
                    break;
                  case 15:
                    sheetName = 'Accounts';
                    return _context7.a(3, 19);
                  case 16:
                    sheetName = 'Flujos';
                    return _context7.a(3, 19);
                  case 17:
                    sheetName = 'Movimientos';
                    return _context7.a(3, 19);
                  case 18:
                    sheetName = downloadType;
                  case 19:
                    _context7.p = 19;
                    existingSheet = context.workbook.worksheets.getItem(sheetName);
                    existingSheet.delete();
                    _context7.n = 20;
                    return context.sync();
                  case 20:
                    console.log("Hoja existente '".concat(sheetName, "' eliminada"));
                    _context7.n = 22;
                    break;
                  case 21:
                    _context7.p = 21;
                    _t9 = _context7.v;
                    // La hoja no existe, no hay problema
                    console.log("La hoja '".concat(sheetName, "' no existe, se crear\xE1 una nueva"));
                  case 22:
                    // Crear la hoja
                    sheet = context.workbook.worksheets.add(sheetName);
                    sheet.load(["protection", "name"]);
                    _context7.n = 23;
                    return context.sync();
                  case 23:
                    if (!sheet.protection.protected) {
                      _context7.n = 24;
                      break;
                    }
                    throw new Error("La hoja está protegida. No se pueden escribir datos.");
                  case 24:
                    console.log("Hoja creada: ".concat(sheetName));

                    // Eliminar Sheet1 si existe (solo la primera vez)
                    _context7.p = 25;
                    sheet1 = context.workbook.worksheets.getItem("Sheet1");
                    sheet1.delete();
                    _context7.n = 26;
                    return context.sync();
                  case 26:
                    console.log("Hoja Sheet1 eliminada");
                    _context7.n = 28;
                    break;
                  case 27:
                    _context7.p = 27;
                    _t0 = _context7.v;
                    // Sheet1 no existe o ya fue eliminada, continuar normalmente
                    console.log("Sheet1 no existe o ya fue eliminada");
                  case 28:
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
                    _context7.n = 29;
                    return context.sync();
                  case 29:
                    showNotification("\xA1".concat(records.length, " ").concat(downloadType, " descargados exitosamente!"), "success");
                  case 30:
                    return _context7.a(2);
                }
              }, _callee7, null, [[25, 27], [19, 21], [6, 9]]);
            }));
            return function (_x2) {
              return _ref3.apply(this, arguments);
            };
          }());
        case 2:
          _context8.n = 4;
          break;
        case 3:
          _context8.p = 3;
          _t1 = _context8.v;
          console.error("Error específico:", _t1.message);
          errorMessage = "Error al descargar los datos"; // Mensajes de error más específicos
          if (_t1.message.includes("protegida")) {
            errorMessage = "La hoja está protegida. Desproteja la hoja e intente nuevamente.";
          } else if (_t1.message.includes("obtener datos")) {
            errorMessage = "Error de conexión. Verifique su conexión a internet.";
          }
          showNotification(errorMessage, "error");
        case 4:
          return _context8.a(2);
      }
    }, _callee8, null, [[1, 3]]);
  }));
  return _download.apply(this, arguments);
}
function importData() {
  return _importData.apply(this, arguments);
}
function _importData() {
  _importData = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee0() {
    var errorMessage, _t12;
    return _regenerator().w(function (_context0) {
      while (1) switch (_context0.p = _context0.n) {
        case 0:
          _context0.p = 0;
          console.log("Iniciando función importData");
          _context0.n = 1;
          return Excel.run(/*#__PURE__*/function () {
            var _ref4 = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee9(context) {
              var application, sheet, range, data, result, retries, response, resultSheetName, counter, resultSheet, resultRange, headerRange, _t10, _t11;
              return _regenerator().w(function (_context9) {
                while (1) switch (_context9.p = _context9.n) {
                  case 0:
                    console.log("Dentro de Excel.run");
                    // Suspender actualización de pantalla
                    application = context.workbook.application;
                    application.suspendScreenUpdatingUntilNextSync();

                    // Validar y obtener datos de origen
                    console.log("Obteniendo hoja activa y rango");
                    sheet = context.workbook.worksheets.getActiveWorksheet();
                    sheet.load("name");
                    range = sheet.getRange("A1:B2");
                    range.load(["values", "rowCount", "columnCount"]);
                    _context9.n = 1;
                    return context.sync();
                  case 1:
                    console.log("Después de sync, valores obtenidos:", range.values);

                    // Validaciones de datos
                    if (!(!range.values || range.values.length < 2)) {
                      _context9.n = 2;
                      break;
                    }
                    console.log("Error: No hay suficientes datos", range.values);
                    throw new Error("No hay suficientes datos para importar");
                  case 2:
                    if (!(!range.values[0][0] || !range.values[0][1])) {
                      _context9.n = 3;
                      break;
                    }
                    console.log("Error: Faltan encabezados", range.values[0]);
                    throw new Error("Los encabezados son requeridos");
                  case 3:
                    if (!(!range.values[1][0] && !range.values[1][1])) {
                      _context9.n = 4;
                      break;
                    }
                    throw new Error("No hay datos para importar");
                  case 4:
                    data = {
                      title: range.values[1][0] || "",
                      body: range.values[1][1] || "",
                      userId: 1
                    };
                    console.log("Preparando datos para enviar:", data);

                    // Intentar enviar datos con reintento
                    retries = 3;
                  case 5:
                    if (!(retries > 0)) {
                      _context9.n = 13;
                      break;
                    }
                    _context9.p = 6;
                    console.log("Intento ".concat(4 - retries, " de env\xEDo de datos"));
                    _context9.n = 7;
                    return fetch('https://jsonplaceholder.typicode.com/posts', {
                      method: 'POST',
                      headers: {
                        'Content-Type': 'application/json'
                      },
                      body: JSON.stringify(data)
                    });
                  case 7:
                    response = _context9.v;
                    if (response.ok) {
                      _context9.n = 8;
                      break;
                    }
                    throw new Error("HTTP error! status: ".concat(response.status));
                  case 8:
                    _context9.n = 9;
                    return response.json();
                  case 9:
                    result = _context9.v;
                    return _context9.a(3, 13);
                  case 10:
                    _context9.p = 10;
                    _t10 = _context9.v;
                    retries--;
                    if (!(retries === 0)) {
                      _context9.n = 11;
                      break;
                    }
                    throw new Error('Error al enviar datos después de 3 intentos');
                  case 11:
                    _context9.n = 12;
                    return new Promise(function (resolve) {
                      return setTimeout(resolve, 1000);
                    });
                  case 12:
                    _context9.n = 5;
                    break;
                  case 13:
                    // Crear hoja de resultado con nombre único
                    resultSheetName = "Resultado";
                    counter = 1;
                  case 14:
                    if (false) // removed by dead control flow
{}
                    _context9.p = 15;
                    context.workbook.worksheets.getItem(resultSheetName);
                    resultSheetName = "Resultado_".concat(counter++);
                    _context9.n = 17;
                    break;
                  case 16:
                    _context9.p = 16;
                    _t11 = _context9.v;
                    return _context9.a(3, 18);
                  case 17:
                    _context9.n = 14;
                    break;
                  case 18:
                    resultSheet = context.workbook.worksheets.add(resultSheetName); // Escribir resultados en un solo bloque
                    resultRange = resultSheet.getRange("A1:C2");
                    resultRange.values = [["ID", "Estado", "Fecha"], [result.id, "Importado exitosamente", new Date().toLocaleString()]];

                    // Formatear la hoja de resultados
                    headerRange = resultRange.getRow(0);
                    headerRange.format.fill.color = "#D3D3D3";
                    headerRange.format.font.bold = true;
                    resultSheet.getUsedRange().format.autofitColumns();
                    _context9.n = 19;
                    return context.sync();
                  case 19:
                    showNotification("¡Datos importados exitosamente!", "success");
                  case 20:
                    return _context9.a(2);
                }
              }, _callee9, null, [[15, 16], [6, 10]]);
            }));
            return function (_x3) {
              return _ref4.apply(this, arguments);
            };
          }());
        case 1:
          _context0.n = 3;
          break;
        case 2:
          _context0.p = 2;
          _t12 = _context0.v;
          console.error("Error específico:", _t12.message);
          errorMessage = "Error al importar los datos"; // Mensajes de error más específicos
          if (_t12.message.includes("suficientes datos")) {
            errorMessage = "No hay suficientes datos para importar. Verifique el rango seleccionado.";
          } else if (_t12.message.includes("enviar datos")) {
            errorMessage = "Error de conexión al enviar datos. Verifique su conexión a internet.";
          } else if (_t12.message.includes("encabezados")) {
            errorMessage = "Los encabezados son requeridos. Verifique la estructura de los datos.";
          }
          showNotification(errorMessage, "error");
        case 3:
          return _context0.a(2);
      }
    }, _callee0, null, [[0, 2]]);
  }));
  return _importData.apply(this, arguments);
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
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\n\n<!DOCTYPE html>\n<html lang=\"es\">\n\n<head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <title>Add-in Azurriga</title>\n\n    <!-- Office JavaScript API -->\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\n\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\n\n    <!-- Template styles -->\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n</head>\n\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\n        <img width=\"90\" height=\"90\" src=\"" + ___HTML_LOADER_IMPORT_1___ + "\" alt=\"Azurriga\" title=\"Azurriga\" />\n         <!-- <h1 class=\"ms-font-su\">Welcome</h1>-->\n    </header>\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\n        <h2 class=\"ms-font-xl\">Please <a target=\"_blank\" rel=\"noopener noreferrer\" href=\"https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins#sideload-an-office-add-in-for-testing\">sideload</a> your add-in to see app body.</h2>\n    </section>\n    <main id=\"app-body\" class=\"ms-welcome__main hidden\">\n        <div class=\"ms-welcome__buttons\">\n            <div role=\"button\" id=\"login\" class=\"ms-Button ms-Button--primary\">\n                <span class=\"ms-Button-label\">Login</span>\n            </div>\n            <div role=\"button\" id=\"download\" class=\"ms-Button ms-Button--primary is-disabled\" disabled>\n                <span class=\"ms-Button-label\">Download</span>\n            </div>\n            <div role=\"button\" id=\"import\" class=\"ms-Button ms-Button--primary is-disabled\" disabled>\n                <span class=\"ms-Button-label\">Import</span>\n            </div>\n        </div>\n        <p><label id=\"item-subject\"></label></p>\n    </main>\n\n    <!-- Popup de Notificación -->\n    <div id=\"notificationPopup\" class=\"notification-popup hidden\">\n        <div class=\"notification-content\">\n            <span id=\"notificationMessage\"></span>\n        </div>\n    </div>\n\n    <!-- Modal de Login -->\n    <div id=\"loginModal\" class=\"modal hidden\">\n        <div class=\"modal-content ms-depth-4\">\n            <h2 class=\"ms-font-xl\">Iniciar Sesión</h2>\n            <div class=\"form-group\">\n                <label class=\"ms-Label\">Usuario:</label>\n                <input type=\"text\" id=\"username\" class=\"ms-TextField-field\" placeholder=\"Ingrese su usuario\">\n            </div>\n            <div class=\"form-group\">\n                <label class=\"ms-Label\">Contraseña:</label>\n                <input type=\"password\" id=\"password\" class=\"ms-TextField-field\" placeholder=\"Ingrese su contraseña\">\n            </div>\n            <div id=\"loginError\" class=\"error-message hidden\"></div>\n            \n            <!-- Loading spinner -->\n            <div id=\"loginLoading\" class=\"loading-container hidden\">\n                <div class=\"spinner\"></div>\n                <p class=\"ms-font-m\">Conectando al servidor...</p>\n            </div>\n            \n            <div class=\"modal-buttons\">\n                <button class=\"ms-Button ms-Button--primary\" id=\"loginSubmit\">\n                    <span class=\"ms-Button-label\">Iniciar Sesión</span>\n                </button>\n                <button class=\"ms-Button\" id=\"loginCancel\">\n                    <span class=\"ms-Button-label\">Cancelar</span>\n                </button>\n            </div>\n        </div>\n    </div>\n\n    <!-- Modal de Descarga -->\n    <div id=\"downloadModal\" class=\"modal hidden\">\n        <div class=\"modal-content ms-depth-4\">\n            <h2 class=\"ms-font-xl\">Opciones de Descarga</h2>\n            \n            <!-- Selector de tipo de descarga -->\n            <div class=\"form-group\">\n                <label class=\"ms-Label\">Tipo de descarga:</label>\n                <select id=\"downloadType\" class=\"ms-Dropdown\" title=\"Seleccionar tipo de descarga\">\n                    <option value=\"cuentas\">Cuentas</option>\n                    <option value=\"flujos\">Flujos</option>\n                    <option value=\"movimientos\">Movimientos</option>\n                </select>\n            </div>\n\n            <!-- Selector de cantidad de registros -->\n            <div class=\"form-group\">\n                <label class=\"ms-Label\">Cantidad de registros:</label>\n                <select id=\"recordLimit\" class=\"ms-Dropdown\" title=\"Seleccionar cantidad de registros\">\n                    <option value=\"50\">50</option>\n                    <option value=\"75\">75</option>\n                    <option value=\"100\">100</option>\n                    <option value=\"500\">500</option>\n                    <option value=\"1000\">1000</option>\n                    <option value=\"all\">Todas</option>\n                </select>\n            </div>\n\n            <!-- Opciones específicas para Movimientos -->\n            <div id=\"movimientosOptions\" class=\"form-group hidden\">\n                <label class=\"ms-Label\">Campos a incluir:</label>\n                <div class=\"checkbox-group\">\n                    <div class=\"ms-CheckBox\">\n                        <input type=\"checkbox\" id=\"fieldStatus\" value=\"Status\" checked>\n                        <label for=\"fieldStatus\">Status</label>\n                    </div>\n                    <div class=\"ms-CheckBox\">\n                        <input type=\"checkbox\" id=\"fieldIsDebit\" value=\"IsDebit\" checked>\n                        <label for=\"fieldIsDebit\">IsDebit</label>\n                    </div>\n                    <div class=\"ms-CheckBox\">\n                        <input type=\"checkbox\" id=\"fieldAmount\" value=\"Amount\" checked>\n                        <label for=\"fieldAmount\">Amount</label>\n                    </div>\n                    <div class=\"ms-CheckBox\">\n                        <input type=\"checkbox\" id=\"fieldValueDate\" value=\"ValueDate\" checked>\n                        <label for=\"fieldValueDate\">ValueDate</label>\n                    </div>\n                    <div class=\"ms-CheckBox\">\n                        <input type=\"checkbox\" id=\"fieldTrnDate\" value=\"TrnDate\" checked>\n                        <label for=\"fieldTrnDate\">TrnDate</label>\n                    </div>\n                    <div class=\"ms-CheckBox\">\n                        <input type=\"checkbox\" id=\"fieldDescription\" value=\"Description\" checked>\n                        <label for=\"fieldDescription\">Description</label>\n                    </div>\n                </div>\n            </div>\n\n            <div id=\"downloadError\" class=\"error-message hidden\"></div>\n            \n            <div class=\"modal-buttons\">\n                <button class=\"ms-Button ms-Button--primary\" id=\"downloadSubmit\">\n                    <span class=\"ms-Button-label\">Descargar</span>\n                </button>\n                <button class=\"ms-Button\" id=\"downloadCancel\">\n                    <span class=\"ms-Button-label\">Cancelar</span>\n                </button>\n            </div>\n        </div>\n    </div>\n</body>\n\n</html>\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map