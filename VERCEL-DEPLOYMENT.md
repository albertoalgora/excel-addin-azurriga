# Vercel OData Proxy

Proxy HTTPS serverless para permitir que Excel Online (HTTPS) se conecte al servidor OData (HTTP) sin problemas de Mixed Content.

## 🚀 Despliegue en Vercel

### Paso 1: Crear cuenta en Vercel (GRATIS)

1. Ve a: https://vercel.com/signup
2. Registrate con tu cuenta de GitHub
3. Autoriza Vercel a acceder a tu repositorio

### Paso 2: Importar proyecto

1. Click en **"Add New..."** → **"Project"**
2. Selecciona el repositorio: **`albertoalgora/excel-addin-azurriga`**
3. Click en **"Import"**

### Paso 3: Configurar proyecto

**Framework Preset:** Other
**Root Directory:** `./` (raíz del proyecto)
**Build Command:** (dejar vacío)
**Output Directory:** (dejar vacío)
**Install Command:** `npm install` (opcional)

Click en **"Deploy"**

### Paso 4: Obtener URL

Una vez desplegado, Vercel te dará una URL como:
```
https://excel-addin-azurriga.vercel.app
```

Tu endpoint del proxy será:
```
https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/AccountSet
```

## 🔧 Uso en el Add-in

El proxy acepta un parámetro `path` en la query string:

```javascript
// Login
fetch('https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/', {
    headers: {
        'Authorization': 'Basic ' + btoa(username + ':' + password)
    }
})

// Descargar cuentas
fetch('https://excel-addin-azurriga.vercel.app/api/proxy?path=odata/AccountSet?$top=50', {
    headers: {
        'Authorization': 'Basic ' + btoa(username + ':' + password)
    }
})
```

## 📊 Características

- ✅ **HTTPS automático** (certificado SSL gratuito)
- ✅ **CORS configurado** para GitHub Pages
- ✅ **Reenvío de Basic Auth** al servidor OData
- ✅ **Soporta todos los métodos HTTP** (GET, POST, PUT, DELETE)
- ✅ **Logs en tiempo real** en Vercel Dashboard
- ✅ **100% gratuito** en Hobby Plan

## 🔐 Seguridad

- Las credenciales NO se almacenan en Vercel
- El proxy solo reenvía las peticiones
- HTTPS end-to-end entre Excel Online y Vercel
- HTTP solo entre Vercel y tu servidor OData (interno)

## 📈 Monitoreo

Dashboard de Vercel muestra:
- Requests por segundo
- Errores y status codes
- Logs en tiempo real
- Uso de bandwidth

## 💰 Costos

**GRATIS** - Hobby Plan incluye:
- Serverless functions ilimitadas
- 100 GB bandwidth/mes
- Sin tarjeta de crédito requerida

## 🔄 Actualización automática

Cada push a GitHub despliega automáticamente en Vercel.

## 🧪 Prueba local

```bash
# Instalar Vercel CLI
npm i -g vercel

# Ejecutar localmente
vercel dev

# Probar
curl "http://localhost:3000/api/proxy?path=odata/" \
  -H "Authorization: Basic dXNlcjpwYXNz"
```

## 🐛 Troubleshooting

### Error: "CORS blocked"
- Ya configurado en `vercel.json`
- Verificar que el header `Authorization` se esté enviando

### Error: "Timeout"
- Máximo 10 segundos por request (configurado en `vercel.json`)
- Optimizar consultas OData si son muy lentas

### Error: "Invalid credentials"
- El proxy reenvía al servidor OData original
- Verificar usuario/contraseña del OData server
