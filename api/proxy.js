/**
 * Vercel Serverless Function - Proxy HTTPS para servidor OData HTTP
 * Resuelve el problema de Mixed Content en Excel Online
 * 
 * URL: https://tu-proyecto.vercel.app/api/proxy?path=odata/AccountSet
 * Version: 2.0 - CORS Fixed
 */

export default async function handler(req, res) {
    // CORS HEADERS - Primero antes que nada
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'X-Requested-With, Content-Type, Authorization, Accept');
    
    console.log(`[Proxy] ${req.method} ${req.url} from ${req.headers.origin || 'no-origin'}`);
    
    // Manejar preflight OPTIONS
    if (req.method === 'OPTIONS') {
        console.log('[Proxy] Handling OPTIONS preflight');
        return res.status(204).end();
    }

    try {
        // Obtener el path de la query string (ej: ?path=odata/AccountSet)
        const { path = '' } = req.query;
        
        // URL del servidor OData original (HTTP)
        const targetUrl = `http://8cf33ac.online-server.cloud:1031/${path}`;
        
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
        try {
            response = await fetch(targetUrl, fetchOptions);
            clearTimeout(timeout);
        } catch (fetchError) {
            clearTimeout(timeout);
            console.error('[Proxy] Fetch error:', fetchError.message);
            
            // Asegurar CORS headers en error
            res.setHeader('Access-Control-Allow-Origin', '*');
            res.setHeader('Access-Control-Allow-Credentials', 'true');
            
            if (fetchError.name === 'AbortError') {
                return res.status(504).json({
                    error: 'Timeout conectando al servidor OData',
                    details: 'El servidor no respondió en 8 segundos',
                    targetUrl: targetUrl
                });
            }
            
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

        // Devolver la respuesta con CORS
        return res.status(response.status).json(data);

    } catch (error) {
        console.error('[Vercel Proxy] Error:', error);
        
        // Asegurar CORS headers en error
        res.setHeader('Access-Control-Allow-Origin', '*');
        res.setHeader('Access-Control-Allow-Credentials', 'true');
        
        // Devolver error con CORS
        return res.status(500).json({
            error: 'Error al conectar con el servidor OData',
            details: error.message,
            timestamp: new Date().toISOString()
        });
    }
}
