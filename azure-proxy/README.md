# Azure OData Proxy Function

Esta Azure Function actúa como proxy HTTPS para permitir que Excel Online (HTTPS) se conecte al servidor OData (HTTP) sin problemas de Mixed Content.

## 🚀 Características

- ✅ Convierte peticiones HTTPS → HTTP
- ✅ Mantiene autenticación Basic Auth
- ✅ CORS configurado para GitHub Pages
- ✅ Soporta todos los métodos HTTP (GET, POST, PUT, DELETE)
- ✅ Reenvía todos los parámetros y headers necesarios

## 📦 Estructura

```
azure-proxy/
├── ODataProxy/
│   ├── function.json    # Configuración del trigger HTTP
│   └── index.js         # Lógica del proxy
├── host.json           # Configuración global de Azure Functions
├── package.json        # Dependencias Node.js
└── README.md          # Esta documentación
```

## 🔧 Instalación Local

```bash
cd azure-proxy
npm install
```

## 🧪 Prueba Local

```bash
# Instalar Azure Functions Core Tools
npm install -g azure-functions-core-tools@4

# Iniciar función localmente
func start
```

La función estará disponible en: `http://localhost:7071/api/{path}`

## ☁️ Despliegue a Azure

### Opción 1: VS Code (Recomendada)

1. Instalar extensión "Azure Functions" en VS Code
2. Click derecho en la carpeta `azure-proxy` → "Deploy to Function App"
3. Seleccionar o crear nueva Function App
4. Esperar despliegue

### Opción 2: Azure CLI

```bash
# Login a Azure
az login

# Crear resource group
az group create --name excel-addin-rg --location westeurope

# Crear storage account
az storage account create --name exceladdinstore --location westeurope --resource-group excel-addin-rg --sku Standard_LRS

# Crear Function App (Consumption Plan = GRATIS)
az functionapp create --resource-group excel-addin-rg --consumption-plan-location westeurope --runtime node --runtime-version 18 --functions-version 4 --name excel-odata-proxy --storage-account exceladdinstore

# Desplegar código
cd azure-proxy
func azure functionapp publish excel-odata-proxy
```

### Opción 3: GitHub Actions (CI/CD)

Ver archivo `.github/workflows/azure-function.yml` (se puede crear si es necesario)

## 🌐 URL de Producción

Después del despliegue, tu URL será:

```
https://<nombre-function-app>.azurewebsites.net/api/{path}
```

Ejemplo:
```
https://excel-odata-proxy.azurewebsites.net/api/odata/AccountSet
```

## 🔐 Seguridad

- La autenticación Basic Auth se reenvía al servidor OData original
- CORS configurado solo para dominios específicos (ajustable)
- Sin almacenamiento de credenciales en Azure
- Tier gratuito incluye hasta 1M de peticiones/mes

## 📊 Monitoreo

- Application Insights incluido automáticamente
- Logs en tiempo real: Azure Portal → Function App → Monitor
- Métricas de uso y rendimiento disponibles

## 💰 Costos

**GRATIS** en tier Consumption:
- 1,000,000 ejecuciones gratis/mes
- 400,000 GB-s de recursos/mes
- Solo pagas si excedes estos límites

Para tu caso de uso → **100% GRATIS**

## 🔄 Actualizar el Add-in

Cambiar las URLs en `src/taskpane/taskpane.js`:

```javascript
// ANTES (HTTP - bloqueado por Mixed Content)
const BASE_URL = 'http://8cf33ac.online-server.cloud:1031/odata';

// DESPUÉS (HTTPS via Azure Function)
const BASE_URL = 'https://excel-odata-proxy.azurewebsites.net/api/odata';
```

## 🐛 Troubleshooting

### Error: "CORS policy blocked"
- Verificar que `Access-Control-Allow-Origin` esté configurado
- Ajustar en `index.js` si necesitas dominios específicos

### Error: "Authentication failed"
- Verificar que el header `Authorization` se esté pasando
- Comprobar credenciales en el servidor OData original

### Error: "Function timeout"
- Aumentar timeout en `host.json` (máx 10 min en Consumption)
- Considerar optimizar consultas OData

## 📝 Notas

- El proxy NO almacena datos, solo reenvía peticiones
- Latencia adicional: ~100-300ms (aceptable para uso normal)
- Compatible con todos los endpoints OData existentes
