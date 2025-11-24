/**
 * Script de desarrollo local para probar el proxy con logging
 * Simula la función serverless de Vercel pero corriendo localmente
 * Ejecutar: node dev-proxy-server.js
 */

import express from 'express';
import cors from 'cors';
import fetch from 'node-fetch';
import { logRequest } from './api/db.js';
import dotenv from 'dotenv';

// Cargar variables de entorno
dotenv.config({ path: '.env.local' });

const app = express();
const PORT = 3002;

// Middleware
app.use(cors());
app.use(express.json());

// Logging middleware
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.path}`);
  next();
});

// Endpoint proxy (igual que api/proxy.js)
app.all('/api/proxy', async (req, res) => {
  const startTime = Date.now();
  
  try {
    const { path = '' } = req.query;
    const targetUrl = `http://8cf33ac.online-server.cloud:1031/${path}`;
    
    console.log(`[Proxy] ${req.method} ${targetUrl}`);
    
    // Preparar headers
    const headers = {
      'Content-Type': req.headers['content-type'] || 'application/json',
      'Accept': 'application/json'
    };
    
    if (req.headers.authorization) {
      headers['Authorization'] = req.headers.authorization;
    }
    
    // Timeout controller
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    
    let response;
    let statusCode;
    let errorMessage = null;
    
    try {
      response = await fetch(targetUrl, {
        method: req.method,
        headers: headers,
        body: req.method !== 'GET' && req.body ? JSON.stringify(req.body) : undefined,
        signal: controller.signal
      });
      
      clearTimeout(timeout);
      statusCode = response.status;
      
    } catch (fetchError) {
      clearTimeout(timeout);
      console.error('[Proxy] Fetch error:', fetchError.message);
      
      if (fetchError.name === 'AbortError') {
        statusCode = 504;
        errorMessage = 'Timeout conectando al servidor OData';
        
        // Guardar log del error
        saveLog(req, path, 504, Date.now() - startTime, errorMessage);
        
        return res.status(504).json({
          error: 'Timeout conectando al servidor OData',
          details: 'El servidor no respondió en 8 segundos',
          targetUrl: targetUrl
        });
      }
      
      statusCode = 502;
      errorMessage = fetchError.message;
      
      saveLog(req, path, 502, Date.now() - startTime, errorMessage);
      
      return res.status(502).json({
        error: 'Error de conexión con el servidor OData',
        details: fetchError.message,
        targetUrl: targetUrl
      });
    }
    
    // Obtener contenido
    const contentType = response.headers.get('content-type');
    let data;
    
    if (contentType && contentType.includes('application/json')) {
      data = await response.json();
    } else {
      data = await response.text();
    }
    
    console.log(`[Proxy] Response: ${statusCode} (${Date.now() - startTime}ms)`);
    
    // Guardar log exitoso
    const responseTime = Date.now() - startTime;
    saveLog(req, path, statusCode, responseTime, null);
    
    return res.status(statusCode).json(data);
    
  } catch (error) {
    console.error('[Proxy] Error:', error);
    
    const responseTime = Date.now() - startTime;
    saveLog(req, path || '', 500, responseTime, error.message);
    
    return res.status(500).json({
      error: 'Error al conectar con el servidor OData',
      details: error.message,
      timestamp: new Date().toISOString()
    });
  }
});

// Endpoint de estadísticas
app.get('/api/stats', async (req, res) => {
  try {
    const { getStats, getAggregatedStats } = await import('./api/db.js');
    const { username, limit = '100', type = 'summary' } = req.query;
    
    const parsedLimit = Math.min(parseInt(limit) || 100, 1000);
    
    if (type === 'detailed') {
      const logs = await getStats(username, parsedLimit);
      return res.json({
        type: 'detailed',
        username: username || 'all',
        count: logs.length,
        limit: parsedLimit,
        logs: logs
      });
    } else {
      const stats = await getAggregatedStats(username);
      return res.json({
        type: 'summary',
        username: username || 'all',
        stats: stats
      });
    }
  } catch (error) {
    console.error('[Stats] Error:', error);
    return res.status(500).json({
      error: 'Error al obtener estadísticas',
      details: error.message
    });
  }
});

/**
 * Mapea el endpoint a un tipo de petición descriptivo
 */
function mapEndpointToTipoPeticion(path) {
  if (!path || path === '' || path === 'odata/' || path === 'odata') {
    return 'Login';
  }
  if (path.includes('AccountSet')) {
    return 'Descarga Cuentas';
  }
  if (path.includes('FlowCodeSet')) {
    return 'Descarga Flujos';
  }
  if (path.includes('CashFlowSet')) {
    return 'Descarga Movimientos';
  }
  // Otros endpoints específicos
  return 'Otro: ' + path;
}

/**
 * Extrae el número de registros solicitados desde el path
 * Busca el parámetro $top en la URL
 */
function extractNumeroRegistros(path, tipoPeticion) {
  // Para login, siempre es 1
  if (tipoPeticion === 'Login') {
    return 1;
  }
  
  // Buscar $top en el path (puede estar codificado como %24top)
  const topMatch = path.match(/[\?&](?:\$|%24)top=(\d+)/i);
  if (topMatch) {
    return parseInt(topMatch[1], 10);
  }
  
  // Si no hay $top, se devuelve null (significa "todos")
  return null;
}

// Función auxiliar para guardar logs
function saveLog(req, path, statusCode, responseTime, errorMessage) {
  let username = 'anonymous';
  try {
    const authHeader = req.headers.authorization || '';
    if (authHeader && authHeader.startsWith('Basic ')) {
      const base64Credentials = authHeader.split(' ')[1];
      const credentials = Buffer.from(base64Credentials, 'base64').toString('utf-8');
      username = credentials.split(':')[0];
    }
  } catch (e) {
    console.error('[Proxy] Error extrayendo username:', e.message);
  }
  
  // Mapear endpoint a tipo de petición
  const tipoPeticion = mapEndpointToTipoPeticion(path);
  
  // Extraer número de registros
  const numeroRegistros = extractNumeroRegistros(path, tipoPeticion);
  
  logRequest({
    username,
    tipoPeticion,
    method: req.method,
    statusCode,
    responseTime,
    userAgent: req.headers['user-agent'] || '',
    errorMessage,
    numeroRegistros
  }).catch(err => {
    console.error('[Proxy] Error guardando log:', err.message);
  });
}

// Iniciar servidor
app.listen(PORT, () => {
  console.log('🚀 Servidor proxy de desarrollo iniciado!');
  console.log(`📍 http://localhost:${PORT}`);
  console.log(`\n📊 Endpoints disponibles:`);
  console.log(`   - Proxy:        http://localhost:${PORT}/api/proxy?path=odata/`);
  console.log(`   - Estadísticas: http://localhost:${PORT}/api/stats`);
  console.log(`\n🔧 Variables de entorno cargadas desde .env.local`);
  console.log(`   - TURSO_DATABASE_URL: ${process.env.TURSO_DATABASE_URL ? '✅ Configurada' : '❌ No encontrada'}`);
  console.log(`   - TURSO_AUTH_TOKEN:   ${process.env.TURSO_AUTH_TOKEN ? '✅ Configurada' : '❌ No encontrada'}`);
  console.log(`\n✨ Listo para recibir peticiones!\n`);
});
