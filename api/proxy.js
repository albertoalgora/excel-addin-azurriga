/**
 * Vercel Serverless Function - Proxy HTTPS para servidor OData HTTP
 * Resuelve el problema de Mixed Content en Excel Online
 * 
 * URL: https://tu-proyecto.vercel.app/api/proxy?path=odata/AccountSet
 * Version: 3.0 - Con logging a Turso DB
 */

import { logRequest } from './db.js';

export default async function handler(req, res) {
    const startTime = Date.now();
    
    // CORS HEADERS - Establecer SIEMPRE primero y para TODOS los casos
    const origin = req.headers.origin || '*';
    
    // Configurar headers CORS antes que nada
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With, Content-Type, Authorization, Accept, Accept-Version, Content-Length');
    res.setHeader('Access-Control-Max-Age', '86400');
    
    console.log(`[Proxy] ${req.method} ${req.url} from ${origin || 'no-origin'}`);
    
    // Manejar preflight OPTIONS - DEBE devolver 200 OK
    if (req.method === 'OPTIONS') {
        console.log('[Proxy] Handling OPTIONS preflight - returning 200 OK');
        res.status(200).end();
        return;
    }

    try {
        // Obtener el path de la query string (ej: ?path=odata/AccountSet)
        const { path = '' } = req.query;
        
        // URL del servidor OData original (HTTP)
        // const targetUrl = `http://8cf33ac.online-server.cloud:1031/${path}`;
        const targetUrl = `https://azprod.azurriga.com:1035/${path}`;
        
        console.log(`[Vercel Proxy] ${req.method} ${targetUrl}`);

        // Preparar headers para el servidor OData
        const headers = {
            'Content-Type': req.headers['content-type'] || 'application/json',
            'Accept': 'application/json'
        };

        // Pasar el header de Authorization si existe (Basic Auth)
        if (req.headers.authorization) {
            headers['Authorization'] = req.headers.authorization;
        }

        // Preparar opciones para fetch con timeout
        const controller = new AbortController();
        const timeout = setTimeout(() => controller.abort(), 8000); // 8 segundos timeout
        
        const fetchOptions = {
            method: req.method,
            headers: headers,
            signal: controller.signal
        };

        // Agregar body si existe (POST, PUT)
        if (req.body && Object.keys(req.body).length > 0) {
            fetchOptions.body = JSON.stringify(req.body);
        }

        // Hacer la petición al servidor OData
        let response;
        let statusCode;
        let errorMessage = null;
        
        try {
            response = await fetch(targetUrl, fetchOptions);
            clearTimeout(timeout);
            statusCode = response.status;
        } catch (fetchError) {
            clearTimeout(timeout);
            console.error('[Proxy] Fetch error:', fetchError.message);
            
            // Determinar status code y error message
            if (fetchError.name === 'AbortError') {
                statusCode = 504;
                errorMessage = 'Timeout conectando al servidor OData';
                
                // Guardar log del error (async, no bloquea respuesta)
                saveLog(req, path, 504, Date.now() - startTime, errorMessage);
                
                return res.status(504).json({
                    error: 'Timeout conectando al servidor OData',
                    details: 'El servidor no respondió en 8 segundos',
                    targetUrl: targetUrl
                });
            }
            
            statusCode = 502;
            errorMessage = fetchError.message;
            
            // Guardar log del error (async, no bloquea respuesta)
            saveLog(req, path, 502, Date.now() - startTime, errorMessage);
            
            return res.status(502).json({
                error: 'Error de conexión con el servidor OData',
                details: fetchError.message,
                targetUrl: targetUrl
            });
        }
        
        // Obtener el contenido de la respuesta
        const contentType = response.headers.get('content-type');
        let data;
        
        if (contentType && contentType.includes('application/json')) {
            data = await response.json();
        } else {
            data = await response.text();
        }

        console.log(`[Vercel Proxy] Response: ${response.status}`);

        // Guardar log de la petición exitosa (async, no bloquea respuesta)
        const responseTime = Date.now() - startTime;
        saveLog(req, path, statusCode, responseTime, null, data);

        // Devolver la respuesta con CORS
        return res.status(response.status).json(data);

    } catch (error) {
        console.error('[Vercel Proxy] Error:', error);
        
        // Guardar log del error (async, no bloquea respuesta)
        const responseTime = Date.now() - startTime;
        saveLog(req, path || '', 500, responseTime, error.message);
        
        // Devolver error con CORS
        return res.status(500).json({
            error: 'Error al conectar con el servidor OData',
            details: error.message,
            timestamp: new Date().toISOString()
        });
    }
}

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
 * Extrae el número de registros reales de la respuesta del API
 * Cuenta los elementos en el array de resultados
 */
function extractNumeroRegistros(tipoPeticion, responseData) {
    // Para login, siempre es 1
    if (tipoPeticion === 'Login') {
        return 1;
    }
    
    // Intentar contar registros en la respuesta
    try {
        // Estructura OData v4: { "@odata.context": "...", "value": [...] }
        if (responseData && Array.isArray(responseData.value)) {
            return responseData.value.length;
        }
        
        // Estructura OData v2: { d: { results: [...] } }
        if (responseData && responseData.d && Array.isArray(responseData.d.results)) {
            return responseData.d.results.length;
        }
        
        // Estructura alternativa: { results: [...] }
        if (responseData && Array.isArray(responseData.results)) {
            return responseData.results.length;
        }
        
        // Si es un array directamente
        if (Array.isArray(responseData)) {
            return responseData.length;
        }
        
        // Si no podemos determinar, null significa "no aplica"
        return null;
    } catch (e) {
        console.error('[Proxy] Error contando registros:', e.message);
        return null;
    }
}

/**
 * Función auxiliar para guardar logs sin bloquear la respuesta
 */
function saveLog(req, path, statusCode, responseTime, errorMessage, responseData = null) {
    // Extraer username del header de Authorization (Basic Auth)
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
    
    // Extraer número de registros de la respuesta real
    const numeroRegistros = extractNumeroRegistros(tipoPeticion, responseData);
    
    // Guardar log (async, no esperar resultado) - Solo si logRequest está disponible
    if (logRequest) {
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
}