/**
 * Script para verificar los logs en la base de datos
 */

import { createClient } from '@libsql/client';
import dotenv from 'dotenv';

dotenv.config({ path: '.env.local' });

async function checkLogs() {
  try {
    console.log('🔍 Verificando logs en Turso...\n');
    
    const db = createClient({
      url: process.env.TURSO_DATABASE_URL,
      authToken: process.env.TURSO_AUTH_TOKEN
    });
    
    // Contar total de registros
    const countResult = await db.execute('SELECT COUNT(*) as total FROM odata_logs');
    console.log(`📊 Total de registros: ${countResult.rows[0].total}\n`);
    
    // Obtener últimos 10 registros
    const logsResult = await db.execute({
      sql: 'SELECT * FROM odata_logs ORDER BY timestamp DESC LIMIT 10',
      args: []
    });
    
    console.log('📋 Últimos 10 registros:\n');
    logsResult.rows.forEach((log, index) => {
      console.log(`${index + 1}. [${log.timestamp}]`);
      console.log(`   Usuario: ${log.username}`);
      console.log(`   Tipo: ${log.tipo_peticion}`);
      console.log(`   Método: ${log.method}`);
      console.log(`   Status: ${log.status_code}`);
      console.log(`   Tiempo: ${log.response_time_ms}ms`);
      console.log('');
    });
    
    // Estadísticas por tipo de petición
    const statsResult = await db.execute(
      'SELECT tipo_peticion, COUNT(*) as count FROM odata_logs GROUP BY tipo_peticion ORDER BY count DESC'
    );
    
    console.log('📈 Resumen por tipo de petición:\n');
    statsResult.rows.forEach(stat => {
      console.log(`   ${stat.tipo_peticion}: ${stat.count} peticiones`);
    });
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

checkLogs();
