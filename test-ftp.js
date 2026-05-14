const ftp = require('basic-ftp');

async function testFTP() {
  const client = new ftp.Client(30000);
  client.ftp.verbose = true;

  const config = {
    host: 'app2-roura-cevasa-com.espacioseguro.com',
    user: 'ia_rc',
    password: 'Roura2026$',
    port: 21,
    secure: true,
    tls: { rejectUnauthorized: false }
  };

  try {
    console.log('🔌 Conectando al servidor FTPS...');
    await client.access(config);
    console.log('✅ Conexión exitosa');

    // Ver directorio actual
    const pwd = await client.pwd();
    console.log('📁 Directorio actual:', pwd);

    // Listar contenido del directorio raíz
    console.log('\n📂 Contenido del directorio raíz:');
    const rootList = await client.list();
    rootList.forEach(item => {
      console.log(`  ${item.type === 2 ? '📁' : '📄'} ${item.name}`);
    });

    // Intentar entrar en /www
    console.log('\n🔍 Intentando entrar en /www...');
    try {
      await client.cd('/www');
      const wwwPwd = await client.pwd();
      console.log('✅ Directorio actual:', wwwPwd);

      const wwwList = await client.list();
      console.log('📂 Contenido de /www:');
      wwwList.forEach(item => {
        console.log(`  ${item.type === 2 ? '📁' : '📄'} ${item.name}`);
      });
    } catch (e) {
      console.log('❌ Error al entrar en /www:', e.message);
    }

    // Verificar si existe el archivo Excel
    console.log('\n🔍 Buscando archivo Excel...');
    try {
      await client.cd('/www');
      const list = await client.list();
      const excelFiles = list.filter(f => f.name.endsWith('.xlsx'));
      if (excelFiles.length > 0) {
        console.log('📊 Archivos Excel encontrados:');
        excelFiles.forEach(f => console.log(`  - ${f.name}`));
      } else {
        console.log('⚠️ No se encontraron archivos .xlsx en /www');
      }
    } catch (e) {
      console.log('❌ Error buscando Excel:', e.message);
    }

  } catch (e) {
    console.error('❌ Error FTP:', e.message);
  } finally {
    client.close();
    console.log('\n👋 Conexión cerrada');
  }
}

testFTP();
