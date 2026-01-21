const fetch = require('node-fetch');

/**
 * Azure Function que actúa como proxy HTTPS hacia servidor OData HTTP
 * Resuelve el problema de Mixed Content en Excel Online
 */
module.exports = async function (context, req) {
    // Configurar CORS para permitir peticiones desde GitHub Pages
    const corsHeaders = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type, Authorization',
        'Access-Control-Max-Age': '86400'
    };

    // Manejar preflight OPTIONS
    if (req.method === 'OPTIONS') {
        context.res = {
            status: 204,
            headers: corsHeaders,
            body: null
        };
        return;
    }

    try {
        // Obtener el path de la petición (ej: /odata/, /odata/AccountSet, etc)
        const path = req.params.path || '';
        
        // URL del servidor OData original (HTTPS)
        const targetUrl = `https://azprod.azurriga.com:1035/${path}`;
        
        // Log para debugging
        context.log(`Proxying request to: ${targetUrl}`);
        context.log(`Method: ${req.method}`);
        context.log(`Headers:`, req.headers);

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

        context.log(`Response status: ${response.status}`);

        // Devolver la respuesta con CORS habilitado
        context.res = {
            status: response.status,
            headers: {
                ...corsHeaders,
                'Content-Type': contentType || 'application/json'
            },
            body: data
        };

    } catch (error) {
        context.log.error('Error en proxy:', error);
        
        context.res = {
            status: 500,
            headers: corsHeaders,
            body: {
                error: 'Error al conectar con el servidor OData',
                details: error.message
            }
        };
    }
};
