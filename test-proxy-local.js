/**
 * Test del proxy con logging a Turso
 * Simula una petición real para verificar que el logging funciona
 */

import 'dotenv/config';
import { logRequest } from './api/db.js';

async function testProxyLogging() {
  console.log('🧪 Probando logging del proxy localmente...\n');
  
  try {
    // Simular datos de una petición real
    const testData = {
      username: 'TEST_USER',
      tipoPeticion: 'Descarga Cuentas',
      method: 'GET',
      statusCode: 200,
      responseTime: 150,
      userAgent: 'Mozilla/5.0 (Test)',
      errorMessage: null,
      numeroRegistros: 42
    };
    
    console.log('📤 Enviando log de prueba:', testData);
    
    await logRequest(testData);
    
    console.log('✅ Log guardado correctamente!\n');
    console.log('🔍 Verifica con: node check-registros.js');
    
  } catch (error) {
    console.error('❌ Error:', error.message);
    console.error(error);
  }
}

testProxyLogging();
