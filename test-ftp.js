const ftp = require('basic-ftp');

async function testFTP() {
  const client = new ftp.Client(30000);
  client.ftp.verbose = true;

  const config = {
    host: 'app2-roura-cevasa-com.espacioseguro.com',
    user: 'app2roura-cevasa',
    password: 'Roura2026$',
    port: 21,
    secure: true,
    tls: { rejectUnauthorized: false }
  };

  try {
    console.log('🔌 Conectando al servidor FTPS...');
    await client.access(config);
    console.log('✅ Conexión exitosa');

    const pwd = await client.pwd();
    console.log('📁 Directorio inicial (pwd):', pwd);

    // Detectar raíz real (mismo algoritmo que ftpConnect en server.js)
    let basePath;
    try {
      await client.cd('/www');
      basePath = '/www';
      console.log('✅ cd(/www) exitoso → BASE_PATH = /www');
    } catch (e) {
      basePath = '/';
      console.log('⚠️  cd(/www) falló → BASE_PATH = / (chroot ya en /www)');
    }

    // Volver a raíz y listar
    await client.cd('/');
    console.log('\n📂 Contenido de la raíz (/):');
    const rootList = await client.list();
    rootList.forEach(item => {
      console.log(`  ${item.type === 2 ? '📁' : '📄'} ${item.name}`);
    });

    // Entrar en basePath y listar
    console.log(`\n📂 Contenido de ${basePath}:`);
    await client.cd(basePath);
    const baseList = await client.list();
    baseList.forEach(item => {
      console.log(`  ${item.type === 2 ? '📁' : '📄'} ${item.name}`);
    });

    // Verificar archivo Excel
    console.log('\n🔍 Buscando archivo Excel (.xlsx) en', basePath + ':');
    const excelFiles = baseList.filter(f => f.name.endsWith('.xlsx'));
    if (excelFiles.length > 0) {
      console.log('📊 Archivos Excel encontrados:');
      excelFiles.forEach(f => console.log(`  - ${f.name}`));
    } else {
      console.log('⚠️  No se encontraron archivos .xlsx en', basePath);
    }

  } catch (e) {
    console.error('❌ Error FTP:', e.message);
  } finally {
    client.close();
    console.log('\n👋 Conexión cerrada');
  }
}

testFTP();
