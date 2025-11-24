/**
 * Script de prueba para verificar conexión con Turso DB
 * Ejecutar: node test-db.js
 */

import { createClient } from '@libsql/client';
import dotenv from 'dotenv';

// Cargar variables de entorno desde .env.local
dotenv.config({ path: '.env.local' });

async function testConnection() {
  console.log('🧪 Probando conexión con Turso DB...\n');
  
  // Verificar que existan las variables
  if (!process.env.TURSO_DATABASE_URL) {
    console.error('❌ Error: TURSO_DATABASE_URL no está configurada en .env.local');
    process.exit(1);
  }
  
  if (!process.env.TURSO_AUTH_TOKEN) {
    console.error('❌ Error: TURSO_AUTH_TOKEN no está configurada en .env.local');
    process.exit(1);
  }
  
  console.log('✅ Variables de entorno encontradas');
  console.log(`📍 URL: ${process.env.TURSO_DATABASE_URL}\n`);
  
  try {
    // Crear cliente
    const client = createClient({
      url: process.env.TURSO_DATABASE_URL,
      authToken: process.env.TURSO_AUTH_TOKEN
    });
    
    console.log('✅ Cliente creado correctamente\n');
    
    // Probar consulta simple
    console.log('📊 Verificando tabla odata_logs...');
    const result = await client.execute('SELECT COUNT(*) as count FROM odata_logs');
    const count = result.rows[0].count;
    
    console.log(`✅ Tabla encontrada! Registros actuales: ${count}\n`);
    
    // Insertar registro de prueba
    console.log('📝 Insertando registro de prueba...');
    await client.execute({
      sql: `INSERT INTO odata_logs (username, endpoint, method, status_code, response_time_ms, user_agent) 
            VALUES (?, ?, ?, ?, ?, ?)`,
      args: ['test_user', 'test/endpoint', 'GET', 200, 123, 'Test Script']
    });
    
    console.log('✅ Registro insertado correctamente\n');
    
    // Verificar el nuevo count
    const result2 = await client.execute('SELECT COUNT(*) as count FROM odata_logs');
    const newCount = result2.rows[0].count;
    
    console.log(`✅ Registros después de insertar: ${newCount}\n`);
    
    // Obtener últimos 5 registros
    console.log('📋 Últimos 5 registros:');
    const latest = await client.execute('SELECT * FROM odata_logs ORDER BY timestamp DESC LIMIT 5');
    
    latest.rows.forEach((row, index) => {
      console.log(`  ${index + 1}. [${row.timestamp}] ${row.username} → ${row.method} ${row.endpoint} (${row.status_code}) - ${row.response_time_ms}ms`);
    });
    
    console.log('\n✅ ¡Todas las pruebas pasaron correctamente! 🎉');
    console.log('La base de datos está lista para usar.\n');
    
  } catch (error) {
    console.error('\n❌ Error:', error.message);
    console.error('\nDetalles:', error);
    process.exit(1);
  }
}

testConnection();
