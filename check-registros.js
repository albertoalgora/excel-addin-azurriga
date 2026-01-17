/**
 * Script para verificar el campo numero_registros en la base de datos
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
    
    const result = await client.execute({
      sql: 'SELECT id, timestamp, username, tipo_peticion, status_code, numero_registros FROM odata_logs ORDER BY id DESC LIMIT 10',
      args: []
    });
    
    console.log(`📊 Últimos 10 registros:\n`);
    
    result.rows.forEach((row, index) => {
      const timestamp = new Date(row.timestamp).toLocaleString('es-ES');
      console.log(`${index + 1}. [${timestamp}]`);
      console.log(`   ID: ${row.id}`);
      console.log(`   Usuario: ${row.username}`);
      console.log(`   Tipo: ${row.tipo_peticion}`);
      console.log(`   Status: ${row.status_code}`);
      console.log(`   Número de registros: ${row.numero_registros === null ? 'NULL' : row.numero_registros}`);
      console.log('');
    });
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

checkRegistros();
