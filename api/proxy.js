/**
 * Vercel Serverless Function - Proxy HTTPS para servidor OData HTTP
 * Resuelve el problema de Mixed Content en Excel Online
 * 
 * URL: https://tu-proyecto.vercel.app/api/proxy?path=odata/AccountSet
 */

export default async function handler(req, res) {
    // Configurar CORS para permitir peticiones desde GitHub Pages y Excel Online
    res.setHeader('Access-Control-Allow-Credentials', 'true');
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization, X-Requested-With');
    
    // Manejar preflight OPTIONS
    if (req.method === 'OPTIONS') {
        res.status(200).end();
        return;
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

        // Preparar opciones para fetch
        const fetchOptions = {
            method: req.method,
            headers: headers
        };

        // Agregar body si existe (POST, PUT)
        if (req.body && Object.keys(req.body).length > 0) {
            fetchOptions.body = JSON.stringify(req.body);
        }

        // Hacer la petición al servidor OData
        const response = await fetch(targetUrl, fetchOptions);
        
        // Obtener el contenido de la respuesta
        const contentType = response.headers.get('content-type');
        let data;
        
        if (contentType && contentType.includes('application/json')) {
            data = await response.json();
        } else {
            data = await response.text();
        }

        console.log(`[Vercel Proxy] Response: ${response.status}`);

        // Devolver la respuesta
        res.status(response.status).json(data);

    } catch (error) {
        console.error('[Vercel Proxy] Error:', error);
        
        res.status(500).json({
            error: 'Error al conectar con el servidor OData',
            details: error.message,
            timestamp: new Date().toISOString()
        });
    }
}
