/**
 * Test directo de la estructura de la base de datos
 */

import { createClient } from '@libsql/client';
import dotenv from 'dotenv';

dotenv.config({ path: '.env.local' });

async function testDbSchema() {
  try {
    console.log('🔍 Verificando estructura de la tabla odata_logs...\n');
    
    const db = createClient({
      url: process.env.TURSO_DATABASE_URL,
      authToken: process.env.TURSO_AUTH_TOKEN
    });
    
    // Obtener estructura de la tabla
    const schema = await db.execute(`PRAGMA table_info(odata_logs)`);
    
    console.log('📋 Columnas de la tabla:');
    schema.rows.forEach(col => {
      console.log(`   - ${col.name} (${col.type})`);
    });
    
    // Verificar si existe la columna tipo_peticion
    const hasTipoPeticion = schema.rows.some(col => col.name === 'tipo_peticion');
    const hasEndpoint = schema.rows.some(col => col.name === 'endpoint');
    
    console.log(`\n✅ tiene 'tipo_peticion': ${hasTipoPeticion}`);
    console.log(`❌ tiene 'endpoint': ${hasEndpoint}`);
    
    if (hasEndpoint) {
      console.log('\n⚠️  La tabla todavía tiene la columna "endpoint"');
      console.log('   Necesitas ejecutar en Turso dashboard:');
      console.log('   ALTER TABLE odata_logs RENAME COLUMN endpoint TO tipo_peticion;');
    } else if (hasTipoPeticion) {
      console.log('\n✅ ¡La tabla está correctamente actualizada con tipo_peticion!');
    } else {
      console.log('\n❌ No se encuentra ni endpoint ni tipo_peticion');
    }
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

testDbSchema();
