/**
 * test-ftp.js — Diagnóstico de conexión FTP
 * Ejecutar con: node test-ftp.js
 */
const ftp  = require('basic-ftp');
const XLSX = require('xlsx');
const { PassThrough } = require('stream');

const EXCEL_FILENAME = 'vw_segplazo_052025.xlsx';

const config = {
  host:     'app2-roura-cevasa-com.espacioseguro.com',
  user:     'app2roura-cevasa',
  password: 'Roura2026$',
  port:     21,
  secure:   true,
  tls:      { rejectUnauthorized: false }
};

async function testFTP() {
  const client = new ftp.Client(30000);
  client.ftp.verbose = true;

  try {
    console.log('\n🔌 Conectando...');
    await client.access(config);
    console.log('✅ Conexión OK');

    const pwd = await client.pwd();
    console.log('📁 Directorio actual tras conectar:', pwd);

    console.log('\n📂 Listado del directorio actual:');
    const list = await client.list();
    if (list.length === 0) {
      console.log('  ⚠️  Lista vacía (posible problema de formato MLSD/LIST)');
    } else {
      list.forEach(f => console.log(`  ${f.type === 2 ? '📁' : '📄'} ${f.name}`));
    }

    console.log(`\n⬇️  Intentando descargar ${EXCEL_FILENAME} directamente...`);
    const pt = new PassThrough();
    const chunks = [];
    pt.on('data', c => chunks.push(c));
    const done = new Promise((ok, fail) => { pt.on('end', ok); pt.on('error', fail); });
    await client.downloadTo(pt, EXCEL_FILENAME);
    await done;
    const buffer = Buffer.concat(chunks);
    console.log(`✅ Descargado: ${buffer.length} bytes`);

    const wb = XLSX.read(buffer, { type: 'buffer' });
    console.log('📊 Hojas del libro:', wb.SheetNames.join(', '));
    const sheet = wb.Sheets['DatosX3'];
    if (sheet) {
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      console.log(`✅ Hoja DatosX3: ${rows.length} filas`);
    } else {
      console.log('❌ Hoja DatosX3 no encontrada');
    }

  } catch (e) {
    console.error('\n❌ Error:', e.message);
  } finally {
    client.close();
    console.log('\n👋 Conexión cerrada');
  }
}

testFTP();
