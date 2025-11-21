/**
 * Vercel Serverless Function - Estadísticas de uso
 * Endpoint: /api/stats?username=juan&limit=50
 * 
 * Parámetros:
 * - username (opcional): Filtrar por usuario específico
 * - limit (opcional): Número de registros (default: 100, max: 1000)
 * - type (opcional): 'detailed' | 'summary' (default: 'summary')
 */

import { getStats, getAggregatedStats } from './db.js';

export default async function handler(req, res) {
    // CORS Headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
    
    // Manejar preflight OPTIONS
    if (req.method === 'OPTIONS') {
        return res.status(204).end();
    }
    
    // Solo permitir GET
    if (req.method !== 'GET') {
        return res.status(405).json({ error: 'Método no permitido' });
    }
    
    try {
        const { username, limit = '100', type = 'summary' } = req.query;
        
        // Validar limit
        const parsedLimit = Math.min(parseInt(limit) || 100, 1000);
        
        if (type === 'detailed') {
            // Retornar logs detallados
            const logs = await getStats(username, parsedLimit);
            
            return res.status(200).json({
                type: 'detailed',
                username: username || 'all',
                count: logs.length,
                limit: parsedLimit,
                logs: logs
            });
        } else {
            // Retornar estadísticas agregadas
            const stats = await getAggregatedStats(username);
            
            return res.status(200).json({
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
}
