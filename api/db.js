/**
 * Turso Database Client
 * Gestiona la conexión y operaciones con la base de datos SQLite
 */

import { createClient } from '@libsql/client';

let dbClient = null;

/**
 * Obtiene o crea la instancia del cliente de base de datos
 */
export function getDbClient() {
  if (!dbClient && process.env.TURSO_DATABASE_URL && process.env.TURSO_AUTH_TOKEN) {
    try {
      dbClient = createClient({
        url: process.env.TURSO_DATABASE_URL,
        authToken: process.env.TURSO_AUTH_TOKEN
      });
      console.log('[DB] Cliente Turso inicializado correctamente');
    } catch (error) {
      console.error('[DB] Error al inicializar cliente:', error.message);
      return null;
    }
  }
  return dbClient;
}

/**
 * Registra una petición OData en la base de datos
 * @param {Object} data - Datos de la petición
 */
export async function logRequest(data) {
  try {
    const db = getDbClient();
    if (!db) {
      console.log('[DB] Base de datos no configurada, saltando log');
      return;
    }
    
    const { username, endpoint, method, statusCode, responseTime, userAgent, errorMessage } = data;
    
    await db.execute({
      sql: `INSERT INTO odata_logs (username, endpoint, method, status_code, response_time_ms, user_agent, error_message) 
            VALUES (?, ?, ?, ?, ?, ?, ?)`,
      args: [
        username || 'anonymous',
        endpoint || '',
        method || 'GET',
        statusCode || 0,
        responseTime || 0,
        userAgent || '',
        errorMessage || null
      ]
    });
    
    console.log(`[DB] Log guardado: ${method} ${endpoint} - ${statusCode} (${responseTime}ms)`);
  } catch (error) {
    // No fallar si el logging falla
    console.error('[DB] Error al guardar log:', error.message);
  }
}

/**
 * Obtiene estadísticas de uso
 * @param {string} username - Usuario específico (opcional)
 * @param {number} limit - Límite de registros (default: 100)
 */
export async function getStats(username = null, limit = 100) {
  try {
    const db = getDbClient();
    if (!db) {
      throw new Error('Base de datos no configurada');
    }
    
    let query, args;
    
    if (username) {
      query = `SELECT * FROM odata_logs WHERE username = ? ORDER BY timestamp DESC LIMIT ?`;
      args = [username, limit];
    } else {
      query = `SELECT * FROM odata_logs ORDER BY timestamp DESC LIMIT ?`;
      args = [limit];
    }
    
    const result = await db.execute({ sql: query, args });
    return result.rows;
  } catch (error) {
    console.error('[DB] Error al obtener estadísticas:', error.message);
    throw error;
  }
}

/**
 * Obtiene estadísticas agregadas
 */
export async function getAggregatedStats(username = null) {
  try {
    const db = getDbClient();
    if (!db) {
      throw new Error('Base de datos no configurada');
    }
    
    const whereClause = username ? `WHERE username = '${username}'` : '';
    
    const queries = [
      // Total de peticiones
      `SELECT COUNT(*) as total FROM odata_logs ${whereClause}`,
      // Peticiones por endpoint
      `SELECT endpoint, COUNT(*) as count FROM odata_logs ${whereClause} GROUP BY endpoint ORDER BY count DESC LIMIT 10`,
      // Peticiones por usuario (solo si no se especifica username)
      username ? null : `SELECT username, COUNT(*) as count FROM odata_logs GROUP BY username ORDER BY count DESC LIMIT 10`,
      // Tiempo de respuesta promedio
      `SELECT AVG(response_time_ms) as avg_response_time FROM odata_logs ${whereClause}`,
      // Errores (status code >= 400)
      `SELECT COUNT(*) as errors FROM odata_logs ${whereClause} AND status_code >= 400`,
      // Peticiones en las últimas 24 horas
      `SELECT COUNT(*) as last_24h FROM odata_logs ${whereClause} AND timestamp >= datetime('now', '-1 day')`
    ].filter(q => q !== null);
    
    const results = await Promise.all(
      queries.map(query => db.execute(query))
    );
    
    return {
      total: results[0].rows[0].total,
      byEndpoint: results[1].rows,
      byUser: username ? null : results[2].rows,
      avgResponseTime: Math.round(results[username ? 2 : 3].rows[0].avg_response_time || 0),
      errors: results[username ? 3 : 4].rows[0].errors,
      last24h: results[username ? 4 : 5].rows[0].last_24h
    };
  } catch (error) {
    console.error('[DB] Error al obtener estadísticas agregadas:', error.message);
    throw error;
  }
}
