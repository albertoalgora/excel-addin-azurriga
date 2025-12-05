/**
 * Script para verificar el campo numero_registros en los logs
 */

import { createClient } from '@libsql/client';
import dotenv from 'dotenv';

dotenv.config({ path: '.env.local' });

const client = createClient({
  url: process.env.TURSO_DATABASE_URL,
  authToken: process.env.TURSO_AUTH_TOKEN
});

async function checkRegistros() {
  try {
    console.log('🔍 Verificando campo numero_registros...\n');
    
    const result = await client.execute(`
      SELECT 
        id,
        datetime(timestamp, 'localtime') as fecha,
        username,
        tipo_peticion,
        status_code,
        numero_registros
      FROM odata_logs
      ORDER BY timestamp DESC
      LIMIT 10
    `);
    
    console.log('📊 Últimos 10 registros con numero_registros:\n');
    
    result.rows.forEach((row, index) => {
      console.log(`${index + 1}. [${row.fecha}]`);
      console.log(`   Usuario: ${row.username}`);
      console.log(`   Tipo: ${row.tipo_peticion}`);
      console.log(`   Status: ${row.status_code}`);
      console.log(`   📊 Número de registros: ${row.numero_registros ?? 'NULL'}`);
      console.log('');
    });
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

checkRegistros();
