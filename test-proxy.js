/**
 * Script rápido para probar el proxy local y verificar que guarda en BD
 */

import fetch from 'node-fetch';

async function testProxy() {
  console.log('🧪 Probando proxy local...\n');
  
  try {
    // Crear credenciales de prueba
    const authString = Buffer.from('usuario_test:password123').toString('base64');
    
    console.log('📤 Enviando petición de login (odata/)...');
    const loginResponse = await fetch('http://localhost:3002/api/proxy?path=odata/', {
      method: 'GET',
      headers: {
        'Authorization': `Basic ${authString}`,
        'Content-Type': 'application/json',
        'User-Agent': 'test-script'
      }
    });
    
    console.log(`✅ Login response: ${loginResponse.status} ${loginResponse.statusText}`);
    
    console.log('\n📤 Enviando petición de descarga cuentas...');
    const accountsResponse = await fetch('http://localhost:3002/api/proxy?path=odata/AccountSet', {
      method: 'GET',
      headers: {
        'Authorization': `Basic ${authString}`,
        'Content-Type': 'application/json',
        'User-Agent': 'test-script'
      }
    });
    
    console.log(`✅ Accounts response: ${accountsResponse.status} ${accountsResponse.statusText}`);
    
    // Esperar 1 segundo para que se guarden los logs
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    console.log('\n📊 Obteniendo estadísticas...');
    const statsResponse = await fetch('http://localhost:3002/api/stats');
    const stats = await statsResponse.json();
    
    console.log('\n📈 Estadísticas:');
    console.log(JSON.stringify(stats, null, 2));
    
  } catch (error) {
    console.error('❌ Error:', error.message);
  }
}

testProxy();
