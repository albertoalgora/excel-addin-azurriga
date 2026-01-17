/**
 * Script para probar el conteo de registros en las descargas
 * Hace peticiones al proxy local y verifica que se guarden los registros correctos
 */

const credentials = Buffer.from('AZURRIGACONS:Azurriga2025').toString('base64');

async function testRequest(path, expectedType) {
  console.log(`\n🧪 Probando: ${expectedType}`);
  console.log(`   Path: ${path}`);
  
  const url = `http://localhost:3002/api/proxy?path=${encodeURIComponent(path)}`;
  
  try {
    const response = await fetch(url, {
      headers: {
        'Authorization': `Basic ${credentials}`
      }
    });
    
    const data = await response.json();
    
    let recordCount = 0;
    if (data && data.d && Array.isArray(data.d.results)) {
      recordCount = data.d.results.length;
    }
    
    console.log(`   ✅ Status: ${response.status}`);
    console.log(`   📊 Registros recibidos: ${recordCount}`);
    
    // Esperar un poco para que se guarde en la BD
    await new Promise(resolve => setTimeout(resolve, 500));
    
  } catch (error) {
    console.log(`   ❌ Error: ${error.message}`);
  }
}

async function runTests() {
  console.log('🚀 Iniciando pruebas de conteo de registros...\n');
  
  // Test 1: Login
  await testRequest('odata/', 'Login');
  
  // Test 2: Descarga limitada de cuentas (5 registros)
  await testRequest('odata/AccountSet?$top=5', 'Descarga Cuentas (5 registros)');
  
  // Test 3: Descarga limitada de flujos (3 registros)
  await testRequest('odata/FlowCodeSet?$top=3', 'Descarga Flujos (3 registros)');
  
  // Test 4: Descarga limitada de movimientos (10 registros)
  await testRequest('odata/CashFlowSet?$top=10', 'Descarga Movimientos (10 registros)');
  
  console.log('\n✅ Pruebas completadas!');
  console.log('📝 Verifica los logs en la base de datos con: node check-logs.js');
}

runTests();
