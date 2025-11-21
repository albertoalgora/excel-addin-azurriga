# Base de Datos - Logging de peticiones OData

Este proyecto usa **Turso** (SQLite distribuido) para registrar todas las peticiones que pasan por el proxy.

## 📊 Estructura de la Base de Datos

### Tabla: `odata_logs`

| Campo | Tipo | Descripción |
|-------|------|-------------|
| `id` | INTEGER | ID único auto-incremental |
| `timestamp` | DATETIME | Fecha y hora de la petición |
| `username` | TEXT | Usuario que realizó la petición |
| `endpoint` | TEXT | Endpoint OData consultado |
| `method` | TEXT | Método HTTP (GET, POST, etc.) |
| `status_code` | INTEGER | Código de estado HTTP |
| `response_time_ms` | INTEGER | Tiempo de respuesta en milisegundos |
| `user_agent` | TEXT | User-Agent del navegador |
| `error_message` | TEXT | Mensaje de error (si hubo) |

## 🚀 Configuración

### 1. Crear base de datos en Turso

```bash
# Opción A: Desde el CLI (si está instalado)
turso db create excel-addin-logs

# Opción B: Desde el Dashboard web
# Ve a https://turso.tech/app y crea la base de datos
```

### 2. Crear la tabla

Ejecuta este SQL en el dashboard de Turso:

```sql
CREATE TABLE odata_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
  username TEXT,
  endpoint TEXT,
  method TEXT,
  status_code INTEGER,
  response_time_ms INTEGER,
  user_agent TEXT,
  error_message TEXT
);

CREATE INDEX idx_timestamp ON odata_logs(timestamp);
CREATE INDEX idx_username ON odata_logs(username);
CREATE INDEX idx_endpoint ON odata_logs(endpoint);
```

### 3. Configurar variables de entorno

**Para desarrollo local:**
```bash
# Copia el archivo de ejemplo
copy .env.example .env.local

# Edita .env.local con tus credenciales de Turso
```

**Para Vercel (producción):**
```bash
vercel env add TURSO_DATABASE_URL
# Pega la URL de tu base de datos

vercel env add TURSO_AUTH_TOKEN
# Pega tu token de autenticación
```

O desde el Dashboard de Vercel:
1. Settings → Environment Variables
2. Agregar `TURSO_DATABASE_URL` y `TURSO_AUTH_TOKEN`

## 📈 Endpoints disponibles

### 1. Proxy con logging automático

**Endpoint:** `https://tu-proyecto.vercel.app/api/proxy?path=odata/AccountSet`

- Todas las peticiones se registran automáticamente
- No requiere cambios en el código del add-in
- El logging es asíncrono (no afecta rendimiento)

### 2. Estadísticas resumidas

**Endpoint:** `https://tu-proyecto.vercel.app/api/stats`

**Parámetros opcionales:**
- `username`: Filtrar por usuario específico
- `type=summary`: Estadísticas agregadas (default)

**Ejemplo:**
```bash
curl https://tu-proyecto.vercel.app/api/stats?username=juan
```

**Respuesta:**
```json
{
  "type": "summary",
  "username": "juan",
  "stats": {
    "total": 1523,
    "byEndpoint": [
      { "endpoint": "odata/AccountSet", "count": 842 },
      { "endpoint": "odata/CashFlowSet", "count": 681 }
    ],
    "byUser": [
      { "username": "juan", "count": 523 },
      { "username": "maria", "count": 1000 }
    ],
    "avgResponseTime": 245,
    "errors": 12,
    "last24h": 89
  }
}
```

### 3. Logs detallados

**Endpoint:** `https://tu-proyecto.vercel.app/api/stats?type=detailed`

**Parámetros opcionales:**
- `username`: Filtrar por usuario
- `limit`: Número de registros (default: 100, max: 1000)

**Ejemplo:**
```bash
curl "https://tu-proyecto.vercel.app/api/stats?type=detailed&username=juan&limit=50"
```

**Respuesta:**
```json
{
  "type": "detailed",
  "username": "juan",
  "count": 50,
  "limit": 50,
  "logs": [
    {
      "id": 1234,
      "timestamp": "2025-11-21T10:30:45.000Z",
      "username": "juan",
      "endpoint": "odata/AccountSet",
      "method": "GET",
      "status_code": 200,
      "response_time_ms": 234,
      "user_agent": "Mozilla/5.0...",
      "error_message": null
    }
  ]
}
```

## 🔍 Consultas SQL útiles

### Top 10 usuarios más activos
```sql
SELECT username, COUNT(*) as requests 
FROM odata_logs 
GROUP BY username 
ORDER BY requests DESC 
LIMIT 10;
```

### Endpoints más lentos
```sql
SELECT endpoint, AVG(response_time_ms) as avg_time 
FROM odata_logs 
WHERE status_code = 200
GROUP BY endpoint 
ORDER BY avg_time DESC 
LIMIT 10;
```

### Errores recientes
```sql
SELECT * FROM odata_logs 
WHERE status_code >= 400 
ORDER BY timestamp DESC 
LIMIT 50;
```

### Actividad por hora
```sql
SELECT 
  strftime('%Y-%m-%d %H:00', timestamp) as hour,
  COUNT(*) as requests
FROM odata_logs 
WHERE timestamp >= datetime('now', '-24 hours')
GROUP BY hour
ORDER BY hour;
```

## 🛠️ Mantenimiento

### Limpiar logs antiguos (opcional)

```sql
-- Eliminar logs de más de 90 días
DELETE FROM odata_logs 
WHERE timestamp < datetime('now', '-90 days');
```

### Ver tamaño de la base de datos

Desde el dashboard de Turso puedes ver:
- Tamaño total de la base de datos
- Número de filas
- Uso de almacenamiento

## 🎯 Plan gratuito de Turso

- **9 GB** de almacenamiento
- **1 billón** de lecturas/mes
- **25 millones** de escrituras/mes
- **3 bases de datos**

Esto es más que suficiente para millones de peticiones al mes.

## 📝 Notas

- El logging es **asíncrono** y no bloquea las respuestas del proxy
- Si falla el logging, la petición continúa normalmente
- Los logs se guardan solo si las variables `TURSO_DATABASE_URL` y `TURSO_AUTH_TOKEN` están configuradas
- Si no están configuradas, el proxy funciona sin logging (útil para desarrollo)
