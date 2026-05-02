const express  = require('express');
const ftp      = require('basic-ftp');
const cors     = require('cors');
const path     = require('path');
const multer   = require('multer');
const PDFDocument = require('pdfkit');
const { Readable, PassThrough } = require('stream');
const bcrypt   = require('bcryptjs');
const jwt      = require('jsonwebtoken');
const fs       = require('fs');
const XLSX     = require('xlsx');
const { execSync } = require('child_process');
const os           = require('os');
const archiver     = require('archiver');

// ── Config ─────────────────────────────────────────────────
const JWT_SECRET = process.env.JWT_SECRET || 'roura-cevasa-secret-2025';
const USERS_FILE = path.join(__dirname, 'data', 'users.json');
const BASE_PATH  = '/www';

// ── Timestamp en formato yyyymmdd_hhmmss para nombres de archivo ──────────
function tsNombre(userId) {
  const n = new Date();
  const p = v => String(v).padStart(2,'0');
  const fecha = `${n.getFullYear()}${p(n.getMonth()+1)}${p(n.getDate())}`;
  const hora  = `${p(n.getHours())}${p(n.getMinutes())}${p(n.getSeconds())}`;
  const base  = `${fecha}_${hora}`;
  return userId ? `${base}_${userId}` : base;
}

const LOGO_PATH  = path.join(__dirname, 'public', 'logo.png');
const FONDO_PATH = path.join(__dirname, 'public', 'fondo_login.jpg');

// Embed logo as base64 for PDF and frontend
const LOGO_B64 = fs.existsSync(LOGO_PATH)
  ? 'data:image/png;base64,' + fs.readFileSync(LOGO_PATH).toString('base64')
  : null;
const FONDO_B64 = fs.existsSync(FONDO_PATH)
  ? 'data:image/jpeg;base64,' + fs.readFileSync(FONDO_PATH).toString('base64')
  : null;

const FTP_CONFIG = {
  host:     process.env.FTP_HOST || 'app2-roura-cevasa-com.espacioseguro.com',
  user:     process.env.FTP_USER || 'ia_rc',
  password: process.env.FTP_PASS || 'Roura2026$',
  port:     parseInt(process.env.FTP_PORT) || 21,
  secure:   true,  // FTPS con SSL
  tls:      { rejectUnauthorized: false }  // Aceptar certificados autofirmados si es necesario
};

// ── Sessions in-memory ──────────────────────────────────────
const activeSessions = {};

// ── User persistence ────────────────────────────────────────
function loadUsers() {
  try {
    if (!fs.existsSync(path.dirname(USERS_FILE))) fs.mkdirSync(path.dirname(USERS_FILE), { recursive: true });
    if (!fs.existsSync(USERS_FILE)) return null;
    return JSON.parse(fs.readFileSync(USERS_FILE, 'utf8'));
  } catch { return null; }
}
function saveUsers(users) {
  fs.mkdirSync(path.dirname(USERS_FILE), { recursive: true });
  fs.writeFileSync(USERS_FILE, JSON.stringify(users, null, 2));
}
function getUsers() {
  let users = loadUsers();
  if (!users) {
    users = [{ id: '1', username: 'admin', password: bcrypt.hashSync('admin123', 10), name: 'Administrador', role: 'admin', createdAt: new Date().toISOString() }];
    saveUsers(users);
  }
  return users;
}

// ── Express ─────────────────────────────────────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 50 * 1024 * 1024,   // 50MB per file
    files: 20,                      // max 20 files per request
    fields: 10,                     // max 10 non-file fields
    fieldSize: 1024 * 1024          // 1MB per field value
  }
});
const app = express();
app.use(cors());
app.use(express.json({ limit: '150mb' }));
app.use(express.urlencoded({ extended: true, limit: '150mb' }));
// Capturar errores de parse JSON y devolver JSON en lugar de HTML
app.use((err, req, res, next) => {
  if (err.type === 'entity.parse.failed' || err.status === 400) {
    return res.status(400).json({ success: false, error: 'Error al parsear el cuerpo de la petición: ' + err.message });
  }
  if (err.status === 413) {
    return res.status(413).json({ success: false, error: 'Archivo demasiado grande. Límite: 150MB.' });
  }
  next(err);
});
app.use(express.static(path.join(__dirname, 'public')));
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

// ── Auth middleware ──────────────────────────────────────────
function requireAuth(req, res, next) {
  const auth = req.headers.authorization;
  if (!auth?.startsWith('Bearer ')) return res.status(401).json({ success: false, error: 'No autenticado' });
  try {
    req.user = jwt.verify(auth.slice(7), JWT_SECRET);
    if (activeSessions[auth.slice(7)]) activeSessions[auth.slice(7)].lastSeen = Date.now();
    next();
  } catch { return res.status(401).json({ success: false, error: 'Sesión expirada' }); }
}
function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.user.role !== 'admin') return res.status(403).json({ success: false, error: 'Acceso restringido a administradores' });
    next();
  });
}
setInterval(() => {
  const cut = Date.now() - 10 * 60 * 1000;
  for (const [t, s] of Object.entries(activeSessions)) if (s.lastSeen < cut) delete activeSessions[t];
}, 60000);

// ── AUTH ────────────────────────────────────────────────────
app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) return res.status(400).json({ success: false, error: 'Faltan credenciales' });
  const users = getUsers();
  const user = users.find(u => u.username === username.trim().toLowerCase());
  if (!user || !bcrypt.compareSync(password, user.password)) return res.status(401).json({ success: false, error: 'Usuario o contraseña incorrectos' });
  const token = jwt.sign({ id: user.id, username: user.username, name: user.name, role: user.role, id_instalador: user.id_instalador || '' }, JWT_SECRET, { expiresIn: '12h' });
  activeSessions[token] = { username: user.username, name: user.name, role: user.role, lat: null, lng: null, accuracy: null, lastSeen: Date.now(), loginAt: Date.now() };
  res.json({ success: true, token, user: { username: user.username, name: user.name, role: user.role } });
});
app.post('/api/auth/logout', requireAuth, (req, res) => { delete activeSessions[req.headers.authorization.slice(7)]; res.json({ success: true }); });
app.post('/api/auth/ping', requireAuth, (req, res) => {
  const token = req.headers.authorization.slice(7);
  const { lat, lng, accuracy } = req.body;
  if (activeSessions[token]) { activeSessions[token].lastSeen = Date.now(); if (lat != null) { activeSessions[token].lat = lat; activeSessions[token].lng = lng; activeSessions[token].accuracy = accuracy; activeSessions[token].geoUpdatedAt = Date.now(); } }
  res.json({ success: true });
});

// ── ADMIN USERS ─────────────────────────────────────────────
app.get('/api/admin/users', requireAdmin, (req, res) => {
  const users = getUsers();
  res.json({ success: true, users: users.map(u => ({
    id: u.id, username: u.username, name: u.name, role: u.role, createdAt: u.createdAt,
    email: u.email || '', apellidos: u.apellidos || '',
    id_delegacion: u.id_delegacion || '', id_instalador: u.id_instalador || '',
    nivel_acceso: u.nivel_acceso || ''
  })) });
});
app.post('/api/admin/users', requireAdmin, (req, res) => {
  const { username, password, name, role, email, apellidos, id_delegacion, id_instalador, nivel_acceso } = req.body;
  if (!username || !password || !name) return res.status(400).json({ success: false, error: 'Faltan campos obligatorios' });
  const users = getUsers();
  if (users.find(u => u.username === username.trim().toLowerCase())) return res.status(409).json({ success: false, error: 'El usuario ya existe' });
  const newUser = {
    id: Date.now().toString(),
    username: username.trim().toLowerCase(),
    password: bcrypt.hashSync(password, 10),
    name: name.trim(),
    role: role === 'admin' ? 'admin' : 'user',
    email: (email || '').trim(),
    apellidos: (apellidos || '').trim(),
    id_delegacion: (id_delegacion || '').trim(),
    id_instalador: (id_instalador || '').trim(),
    nivel_acceso: (nivel_acceso || '').trim(),
    createdAt: new Date().toISOString()
  };
  users.push(newUser); saveUsers(users);
  res.json({ success: true, user: { id: newUser.id, username: newUser.username, name: newUser.name, role: newUser.role } });
});
app.put('/api/admin/users/:id', requireAdmin, (req, res) => {
  const users = getUsers();
  const idx = users.findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ success: false, error: 'Usuario no encontrado' });
  const { name, email, apellidos, id_delegacion, id_instalador, nivel_acceso, role, password } = req.body;
  if (name) users[idx].name = name.trim();
  if (email !== undefined) users[idx].email = email.trim();
  if (apellidos !== undefined) users[idx].apellidos = apellidos.trim();
  if (id_delegacion !== undefined) users[idx].id_delegacion = id_delegacion.trim();
  if (id_instalador !== undefined) users[idx].id_instalador = id_instalador.trim();
  if (nivel_acceso !== undefined) users[idx].nivel_acceso = nivel_acceso.trim();
  if (role) users[idx].role = role === 'admin' ? 'admin' : 'user';
  if (password && password.length >= 4) users[idx].password = bcrypt.hashSync(password, 10);
  saveUsers(users);
  res.json({ success: true });
});
app.delete('/api/admin/users/:id', requireAdmin, (req, res) => {
  const users = getUsers();
  const idx = users.findIndex(u => u.id === req.params.id);
  if (idx === -1) return res.status(404).json({ success: false, error: 'Usuario no encontrado' });
  if (users[idx].username === 'admin') return res.status(403).json({ success: false, error: 'No se puede eliminar el administrador principal' });
  const del = users[idx].username;
  for (const [t, s] of Object.entries(activeSessions)) if (s.username === del) delete activeSessions[t];
  users.splice(idx, 1); saveUsers(users);
  res.json({ success: true });
});

// ── EXPORT / IMPORT USERS (Excel) ───────────────────────────
app.get('/api/admin/users/export-excel', requireAdmin, (req, res) => {
  try {
    const users = getUsers();
    const rows = users.map(u => ({
      'Id_Usuario':        u.id || '',
      'Email':             u.email || '',
      'Nombre':            u.name || '',
      'Apellidos':         u.apellidos || '',
      'Id_Delegacion':     u.id_delegacion || '',
      'Id_Instalador':     u.id_instalador || '',
      'Usuario_Password':  u.username || '',
      'Nivel_Acceso':      u.nivel_acceso || '',
      'Rol_Sistema':       u.role || 'user',
      'Fecha_Creacion':    u.createdAt || ''
    }));
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    // Column widths
    ws['!cols'] = [20,30,20,20,18,18,20,18,14,22].map(w=>({wch:w}));
    XLSX.utils.book_append_sheet(wb, ws, 'Usuarios');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="usuarios_export.xlsx"');
    res.send(buf);
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/admin/users/import-excel', requireAdmin, upload.single('file'), (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, error: 'Sin archivo' });
    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const users = getUsers();
    let created = 0, updated = 0;
    for (const r of rows) {
      const username = String(r['Usuario_Password'] || '').trim().toLowerCase();
      if (!username) continue;
      const existing = users.find(u => u.username === username);
      if (existing) {
        if (r['Nombre'])       existing.name          = String(r['Nombre']).trim();
        if (r['Apellidos'])    existing.apellidos     = String(r['Apellidos']).trim();
        if (r['Email'])        existing.email         = String(r['Email']).trim();
        if (r['Id_Delegacion'])existing.id_delegacion = String(r['Id_Delegacion']).trim();
        if (r['Id_Instalador'])existing.id_instalador = String(r['Id_Instalador']).trim();
        if (r['Nivel_Acceso']) existing.nivel_acceso  = String(r['Nivel_Acceso']).trim();
        if (r['Rol_Sistema'])  existing.role          = r['Rol_Sistema']==='admin'?'admin':'user';
        updated++;
      } else {
        const pwd = String(r['Contrasena'] || r['Contraseña'] || '').trim() || username;
        users.push({
          id: String(r['Id_Usuario'] || Date.now()).trim() || Date.now().toString(),
          username,
          password: bcrypt.hashSync(pwd, 10),
          name: String(r['Nombre'] || username).trim(),
          apellidos: String(r['Apellidos'] || '').trim(),
          email: String(r['Email'] || '').trim(),
          id_delegacion: String(r['Id_Delegacion'] || '').trim(),
          id_instalador: String(r['Id_Instalador'] || '').trim(),
          nivel_acceso: String(r['Nivel_Acceso'] || '').trim(),
          role: r['Rol_Sistema'] === 'admin' ? 'admin' : 'user',
          createdAt: new Date().toISOString()
        });
        created++;
      }
    }
    saveUsers(users);
    res.json({ success: true, created, updated });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.get('/api/admin/users/export', requireAdmin, (req, res) => {
  try {
    const raw = fs.existsSync(USERS_FILE) ? fs.readFileSync(USERS_FILE, 'utf8') : '[]';
    res.setHeader('Content-Type', 'application/json');
    res.setHeader('Content-Disposition', 'attachment; filename="usuarios_backup.json"');
    res.send(raw);
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

app.post('/api/admin/users/import', requireAdmin, (req, res) => {
  try {
    const imported = req.body.users;
    if (!Array.isArray(imported)) return res.status(400).json({ success: false, error: 'Formato inválido' });
    for (const u of imported) {
      if (!u.id || !u.username || !u.password || !u.name) return res.status(400).json({ success: false, error: 'Registros incompletos en el fichero' });
    }
    saveUsers(imported);
    res.json({ success: true, count: imported.length });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
});

// ── LOCATIONS ───────────────────────────────────────────────
app.get('/api/admin/locations', requireAdmin, (req, res) => {
  const now = Date.now();
  const locs = Object.values(activeSessions).filter(s => s.lastSeen > now - 10 * 60 * 1000).map(s => ({ username: s.username, name: s.name, role: s.role, lat: s.lat, lng: s.lng, accuracy: s.accuracy, lastSeen: s.lastSeen, geoUpdatedAt: s.geoUpdatedAt || null, loginAt: s.loginAt, hasLocation: s.lat !== null }));
  res.json({ success: true, locations: locs, timestamp: now });
});

// ── ZIP HELPER ───────────────────────────────────────────────
/**
 * Crea un buffer ZIP en memoria a partir de un array de entradas:
 *   entries = [{ name: 'fichero.jpg', buffer: Buffer }, ...]
 */
function crearZipBuffer(entries) {
  return new Promise((resolve, reject) => {
    const archive = archiver('zip', { zlib: { level: 6 } });
    const chunks  = [];
    archive.on('data',    c  => chunks.push(c));
    archive.on('end',     () => resolve(Buffer.concat(chunks)));
    archive.on('error',   e  => reject(e));
    for (const { name, buffer } of entries) {
      archive.append(Readable.from(buffer), { name });
    }
    archive.finalize();
  });
}

// ── FTP HELPERS ─────────────────────────────────────────────
async function ensureDir(client, remotePath) {
  // Try direct cd first
  try { await client.cd(remotePath); return; } catch (_) {}

  // Build path segment by segment
  const parts = remotePath.replace(/^\/+/, '').split('/').filter(Boolean);
  let current = '';
  for (const part of parts) {
    current += '/' + part;
    try {
      await client.cd(current);
    } catch (_) {
      // Directory doesn't exist — create it
      try {
        await client.send('MKD ' + current);
      } catch (mkdErr) {
        // 550 = already exists (race condition), 521 = already exists on some servers
        const msg = mkdErr.message || '';
        if (!msg.includes('550') && !msg.includes('521') && !msg.includes('File exists')) {
          throw new Error('MKD failed for ' + current + ': ' + msg);
        }
      }
      // Now try cd again after mkdir
      try {
        await client.cd(current);
      } catch (cdErr) {
        throw new Error('Cannot enter ' + current + ': ' + cdErr.message);
      }
    }
  }
}

// ── FTP ENDPOINTS ────────────────────────────────────────────

// ── Servir logo.js con variables base64 para el frontend ─────────────────
app.get('/logo.js', (req, res) => {
  res.setHeader('Content-Type', 'application/javascript');
  res.send(
    'const LOGO_B64 = ' + JSON.stringify(LOGO_B64 || null) + ';\n' +
    'const FONDO_B64 = ' + JSON.stringify(FONDO_B64 || null) + ';\n'
  );
});

app.get('/api/list', requireAuth, async (req, res) => {
  const dirPath = req.query.path || '/';
  const client = new ftp.Client(30000); client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);
    await client.cd(dirPath === '/' ? BASE_PATH : path.posix.join(BASE_PATH, dirPath));
    const list = await client.list();
    res.json({ success: true, path: dirPath, items: list.map(i => ({ name: i.name, type: i.type === 2 ? 'directory' : 'file', size: i.size, modifiedAt: i.modifiedAt })) });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
  finally { client.close(); }
});

app.get('/api/download', requireAuth, async (req, res) => {
  const filePath = req.query.path;
  if (!filePath) return res.status(400).json({ error: 'Falta path' });
  const client = new ftp.Client(30000); client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);
    res.setHeader('Content-Disposition', `attachment; filename="${path.basename(filePath)}"`);
    await client.downloadTo(res, path.posix.join(BASE_PATH, filePath));
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
  finally { client.close(); }
});

app.post('/api/upload', requireAuth, upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, error: 'Sin archivo' });
  const client = new ftp.Client(60000); client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);
    const tDir = req.body.path === '/' ? BASE_PATH : path.posix.join(BASE_PATH, req.body.path || '/');
    await ensureDir(client, tDir);
    const pedido_upload = (req.body.pedido || '0').toString().trim();
    const ts_upload = tsNombre(req.user.id || '');
    const ext_upload = path.extname(req.file.originalname);
    const base_upload = path.basename(req.file.originalname, ext_upload).replace(/[^a-zA-Z0-9_\-.]/g, '_').slice(0, 80);
    const newName = pedido_upload + '_' + ts_upload + '_' + base_upload + ext_upload;
    await client.uploadFrom(Readable.from(req.file.buffer), path.posix.join(tDir, newName));
    res.json({ success: true, fileName: newName });
  } catch (e) { res.status(500).json({ success: false, error: e.message }); }
  finally { client.close(); }
});

app.post('/api/upload-batch', requireAuth, upload.array('files', 20), async (req, res) => {
  const { pedido, categoria } = req.body;

  console.log('[upload-batch] pedido=%s categoria=%s files=%d', pedido, categoria, req.files?.length || 0);
  if (!pedido) return res.status(400).json({ success: false, error: 'Falta número de pedido. Campos recibidos: ' + JSON.stringify(Object.keys(req.body)) });
  if (!req.files || req.files.length === 0) return res.status(400).json({ success: false, error: 'No se recibieron archivos. Verifica que el campo se llame "files".' });

  // Carpeta de destino según categoría — se intentan múltiples variantes de nombre
  const categoriaMap = {
    'fotos_medicion':    { dirs: ['TD'],                                                              label: 'Fotografias_Medicion' },
    'fotos_antes':       { dirs: ['Fotos Inicio','Fotos inicio','fotos inicio','FOTOS INICIO'],    label: 'Fotografias_Antes' },
    'fotos_final':       { dirs: ['Fotos Fin','Fotos fin','fotos fin','FOTOS FIN'],                label: 'Fotografias_Final' },
    'fotos_cfo':         { dirs: ['Fotos Fin','Fotos fin','fotos fin','FOTOS FIN'],                label: 'Fotografias Visita CFO' },
    'fotos_cierre_cfo':  { dirs: ['Fotos Fin','Fotos fin','fotos fin','FOTOS FIN'],                label: 'Fotografias Cierre Objeciones CFO' },
    'fotos_expediciones':{ dirs: ['Expediciones'],                                                  label: 'Fotografias Expediciones' },
    'fotos_inspecciones':{ dirs: ['Inspecciones Seguridad'],                                        label: 'Fotografias Inspeccion Seguridad' },
  };
  const cat     = String(categoria || 'otros').trim();
  const catInfo = categoriaMap[cat] || { dirs: [cat], label: cat };

  const client = new ftp.Client(180000);
  client.ftp.verbose = false;

  try {
    // ── Construir entradas del ZIP ────────────────────────
    const entries = req.files.map(file => {
      const ext  = path.extname(file.originalname) || '';
      const base = path.basename(file.originalname, ext)
                       .replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 60);
      const safe = Date.now() + '_' + base + ext;
      return { name: safe, buffer: file.buffer, original: file.originalname, size: file.size };
    });

    // ── Comprimir en un único ZIP en memoria ──────────────
    const zipBuffer  = await crearZipBuffer(entries.map(e => ({ name: e.name, buffer: e.buffer })));
    const ts         = tsNombre(req.user.id || '');
    const zipName    = `${String(pedido).trim()}_${ts}_${catInfo.label}.zip`;

    // ── Subir el ZIP al FTP probando variantes de directorio ─
    await client.access(FTP_CONFIG);

    // Intentar cada variante del nombre de directorio hasta que una funcione
    let uploadedOk = false;
    let lastErr    = null;
    for (const dirName of catInfo.dirs) {
      try {
        const tDir = path.posix.join(BASE_PATH, dirName);
        await ensureDir(client, tDir);
        await client.uploadFrom(Readable.from(zipBuffer), path.posix.join(tDir, zipName));
        uploadedOk = true;
        break;
      } catch (e) {
        lastErr = e;
        // Reset client position before trying next variant
        try { await client.cd(BASE_PATH); } catch(_) {}
      }
    }
    if (!uploadedOk) { console.error('[upload-batch] Todos los dirs fallaron:', catInfo.dirs, lastErr?.message); throw lastErr || new Error('No se pudo subir a ningún directorio'); }

    // ── Si son fotos antes/final, copiar también a "Fotografias Pte Certificacion" ──
    if (cat === 'fotos_antes' || cat === 'fotos_final') {
      try {
        const dirPteCert = path.posix.join(BASE_PATH, 'Fotografias Pte Certificacion');
        await ensureDir(client, dirPteCert);
        await client.uploadFrom(Readable.from(zipBuffer), path.posix.join(dirPteCert, zipName));
        console.log('[upload-batch] Copia en Fotografias Pte Certificacion:', zipName);
      } catch (eCert) {
        console.warn('[upload-batch] No se pudo copiar a Fotografias Pte Certificacion:', eCert.message);
      }
    }

    res.json({
      success: true,
      count:   entries.length,
      zipFile: zipName,
      uploaded: entries.map(e => ({ original: e.original, saved: e.name, size: e.size })),
      errors:   []
    });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  } finally {
    client.close();
  }
});

// ── ALBARAN ──────────────────────────────────────────────────
app.post('/api/albaran', requireAuth, async (req, res) => {
  const { pedido, tipo, observaciones, clienteNombre, clienteDni, firma, timestamp, geoData, datosExcel } = req.body;
  if (!pedido) return res.status(400).json({ success: false, error: 'Falta número de pedido' });
  let pdfBuffer;
  try { pdfBuffer = await generarPDF({ pedido, tipo, observaciones, clienteNombre, clienteDni, firma, timestamp, geoData, operario: req.user.name }); }
  catch (e) { return res.status(500).json({ success: false, error: 'Error generando PDF: ' + e.message }); }

  const client = new ftp.Client(300000); client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);
    const albDirMap = {
      'medicion':    { dir: 'Albaranes TD',            label: 'Albaran Medicion' },
      'instalacion': { dir: 'Albaranes Final Trabajo',  label: 'Albaran Final Trabajo' },
    };
    const albInfo = albDirMap[tipo] || { dir: 'Albaranes TD', label: 'Albaran' };
    const tDir    = path.posix.join(BASE_PATH, albInfo.dir);
    await ensureDir(client, tDir);
    const tsAlb   = tsNombre(req.user.id || '');
    const pdfName = `${String(pedido).trim()}_${tsAlb}_${albInfo.label}.pdf`;
    await client.uploadFrom(Readable.from(pdfBuffer), path.posix.join(tDir, pdfName));

    // ── Si es albarán final de instalación, generar Certificacion Final Trabajo ──
    if (tipo === 'instalacion') {
      // Usamos un cliente FTP independiente para toda la lógica de certificación,
      // evitando que los cd() de ensureDir corrompan el estado del cliente principal.
      const clientCert = new ftp.Client(300000);
      clientCert.ftp.verbose = false;
      try {
        await clientCert.access(FTP_CONFIG);

        const dirPteCert = path.posix.join(BASE_PATH, 'Fotografias Pte Certificacion');
        const dirCert    = path.posix.join(BASE_PATH, 'Certificacion Final Trabajo');
        const pedidoStr  = String(pedido).trim();

        // ── 1. Subir el albarán a "Fotografias Pte Certificacion" como ZIP ──────
        try {
          const albZipBuf  = await crearZipBuffer([{ name: pdfName, buffer: pdfBuffer }]);
          const albZipName = `${pedidoStr}_${tsNombre(req.user.id || '')}_Albaran_Final_Trabajo.zip`;
          await clientCert.cd(BASE_PATH);
          await ensureDir(clientCert, dirPteCert);
          await clientCert.uploadFrom(Readable.from(albZipBuf), path.posix.join(dirPteCert, albZipName));
          console.log('[cert] Albarán copiado a Fotografias Pte Certificacion:', albZipName);
        } catch (eAlbCopy) {
          console.warn('[cert] No se pudo copiar albarán a Fotografias Pte Certificacion:', eAlbCopy.message);
        }

        // ── 2. Listar ZIPs del pedido en Fotografias Pte Certificacion ──────────
        let listPte = [];
        try {
          await clientCert.cd(dirPteCert);
          listPte = await clientCert.list();
        } catch (eList) {
          console.warn('[cert] No se pudo listar Fotografias Pte Certificacion:', eList.message);
        }
        const zipsDePedido = listPte.filter(f =>
          f.type !== 2 &&
          /\.zip$/i.test(f.name) &&
          (f.name.startsWith(pedidoStr + '_') || f.name.startsWith(pedidoStr + ' '))
        );
        console.log('[cert] ZIPs encontrados para pedido', pedidoStr, ':', zipsDePedido.map(z => z.name));

        // ── 3. Descargar y clasificar los ZIPs ──────────────────────────────────
        const fotosAntesBufs = [];
        const fotosFinalBufs = [];
        const albaranBuf     = pdfBuffer; // ya disponible en memoria

        for (const zipFile of zipsDePedido) {
          // El albarán ya lo tenemos en pdfBuffer, no hace falta descargarlo
          if (/Albaran_Final_Trabajo/i.test(zipFile.name)) {
            console.log('[cert] Albarán ZIP omitido (usando buffer en memoria):', zipFile.name);
            continue;
          }
          try {
            // Posicionarse en la carpeta antes de cada descarga
            await clientCert.cd(dirPteCert);
            const pt2     = new PassThrough();
            const chunks2 = [];
            pt2.on('data', c => chunks2.push(c));
            const done2 = new Promise((res2, rej2) => { pt2.on('end', res2); pt2.on('error', rej2); });
            await clientCert.downloadTo(pt2, path.posix.join(dirPteCert, zipFile.name));
            await done2;
            const zipBuf = Buffer.concat(chunks2);

            // Extraer en directorio temporal
            const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cert_'));
            try {
              const tmpZip = path.join(tmpDir, 'fotos.zip');
              fs.writeFileSync(tmpZip, zipBuf);
              execSync(`unzip -o "${tmpZip}" -d "${tmpDir}"`, { timeout: 60000 });
              const esAntes = /Fotografias_Antes/i.test(zipFile.name);
              const allFiles = fs.readdirSync(tmpDir).filter(fn => fn !== 'fotos.zip').sort();
              for (const fn of allFiles) {
                if (!esImagen(fn)) continue;
                const buf = fs.readFileSync(path.join(tmpDir, fn));
                if (esAntes) fotosAntesBufs.push({ name: fn, buf });
                else         fotosFinalBufs.push({ name: fn, buf });
              }
              console.log('[cert] ZIP', zipFile.name, '->', allFiles.filter(esImagen).length, 'imgs (', esAntes ? 'ANTES' : 'FINAL', ')');
            } finally {
              try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (_) {}
            }
          } catch (eZip) {
            console.warn('[cert] Error extrayendo ZIP', zipFile.name, ':', eZip.message);
          }
        }
        console.log('[cert] Fotos antes:', fotosAntesBufs.length, '| Fotos final:', fotosFinalBufs.length, '| Albarán OK:', !!albaranBuf);

        // ── 4. Generar PDF de certificación ──────────────────────────────────────
        const fechaHora  = timestamp || new Date().toLocaleString('es-ES');
        const pdfCertBuf = await generarCertificacion({
          pedido: pedidoStr, datosExcel: datosExcel || {},
          fechaHora, operario: req.user.name,
          fotosAntesBufs, fotosFinalBufs, albaranBuf
        });

        // ── 5. Subir la certificación ─────────────────────────────────────────────
        await clientCert.cd(BASE_PATH);
        await ensureDir(clientCert, dirCert);
        const tsCert   = tsNombre(req.user.id || '');
        const certName = `${pedidoStr}_${tsCert}_Certificacion_Final_Trabajo.pdf`;
        await clientCert.uploadFrom(Readable.from(pdfCertBuf), path.posix.join(dirCert, certName));
        console.log('[cert] Certificación subida:', certName);

        // ── 6. Eliminar todos los ZIPs del pedido de Fotografias Pte Certificacion ─
        try {
          await clientCert.cd(dirPteCert);
          const listFinal = await clientCert.list();
          const aEliminar = listFinal.filter(f =>
            f.type !== 2 &&
            /\.zip$/i.test(f.name) &&
            (f.name.startsWith(pedidoStr + '_') || f.name.startsWith(pedidoStr + ' '))
          );
          for (const zipFile of aEliminar) {
            try {
              await clientCert.remove(path.posix.join(dirPteCert, zipFile.name));
              console.log('[cert] Eliminado:', zipFile.name);
            } catch (eDel) {
              console.warn('[cert] No se pudo eliminar', zipFile.name, ':', eDel.message);
            }
          }
        } catch (eClean) {
          console.warn('[cert] Error en limpieza de ZIPs:', eClean.message);
        }

        return res.json({
          success: true, fileName: pdfName, ruta: `${tDir}/${pdfName}`,
          certificacion: { fileName: certName, ruta: `${dirCert}/${certName}` }
        });
      } catch (eCert) {
        console.error('[cert] Error generando certificación automática:', eCert.message);
        return res.json({
          success: true, fileName: pdfName, ruta: `${tDir}/${pdfName}`,
          certificacionError: 'Albarán subido correctamente, pero falló la generación automática de la certificación: ' + eCert.message
        });
      } finally {
        clientCert.close();
      }
    }

    res.json({ success: true, fileName: pdfName, ruta: `${tDir}/${pdfName}` });
  } catch (e) { res.status(500).json({ success: false, error: 'Error subiendo albarán: ' + e.message }); }
  finally { client.close(); }
});

// ── PDF GENERATOR ─────────────────────────────────────────────
function generarPDF({ pedido, tipo, observaciones, clienteNombre, clienteDni, firma, timestamp, geoData, operario }) {
  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 0, size: 'A4' });
    const bufs = [];
    doc.on('data', c => bufs.push(c));
    doc.on('end', () => resolve(Buffer.concat(bufs)));
    doc.on('error', reject);

    const W = doc.page.width;   // 595
    const H = doc.page.height;  // 842
    const M = 40; // margin
    const CW = W - M * 2;       // content width

    const tipoLabel = tipo === 'medicion' ? 'TOMA DE DATOS / MEDICIÓN' : 'DOCUMENTACIÓN DE INSTALACIÓN';
    const tipoSub   = tipo === 'medicion' ? 'Albarán de Medición' : 'Albarán de Instalación';
    const fechaHora = timestamp || new Date().toLocaleString('es-ES');
    const BLUE      = '#1B6CA8'; // R&C brand blue
    const DARK      = '#1a2332';
    const LGRAY     = '#f4f6f9';
    const BORDER    = '#d0d8e4';
    const TEXTGRAY  = '#5a6478';

    // ── HEADER BAR ──────────────────────────────────────────
    doc.rect(0, 0, W, 110).fill(DARK);

    // Logo area (white rounded bg)
    doc.roundedRect(M, 18, 90, 74, 8).fill('#ffffff');
    // Try to embed logo
    const logoB64 = LOGO_B64;
    if (logoB64) {
      try {
        const b64data = logoB64.replace(/^data:image\/\w+;base64,/, '');
        const logoBuf = Buffer.from(b64data, 'base64');
        doc.image(logoBuf, M + 5, 23, { width: 80, height: 64, fit: [80, 64], align: 'center', valign: 'center' });
      } catch (_) {}
    }

    // Company name & doc type
    doc.font('Helvetica-Bold').fontSize(18).fillColor('#ffffff')
       .text('ROURA & CEVASA', M + 102, 24, { width: 260 });
    doc.font('Helvetica').fontSize(10).fillColor('#a8c4e0')
       .text('Gestión Técnica y Documentación', M + 102, 46, { width: 260 });
    doc.font('Helvetica').fontSize(9).fillColor('#7a9abf')
       .text('82.98.168.246 · roura-cevasa', M + 102, 62, { width: 260 });

    // Doc type badge on right
    doc.roundedRect(W - M - 150, 22, 150, 66, 6).fill(BLUE);
    doc.font('Helvetica-Bold').fontSize(8).fillColor('#cde4f5')
       .text('DOCUMENTO', W - M - 150, 32, { width: 150, align: 'center' });
    doc.font('Helvetica-Bold').fontSize(11).fillColor('#ffffff')
       .text(tipoSub.toUpperCase(), W - M - 150, 46, { width: 150, align: 'center' });
    doc.font('Helvetica').fontSize(8).fillColor('#cde4f5')
       .text(fechaHora, W - M - 150, 68, { width: 150, align: 'center' });

    // Blue accent strip
    doc.rect(0, 110, W, 6).fill(BLUE);

    let y = 128;

    // ── HELPER FUNCTIONS ─────────────────────────────────────
    const sectionHeader = (title, icon, yPos) => {
      doc.rect(M, yPos, CW, 24).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(10).fillColor('#ffffff')
         .text(`${icon}  ${title}`, M + 10, yPos + 7, { width: CW - 20 });
      return yPos + 24;
    };

    const fieldRow = (label, value, xPos, yPos, w, h = 32) => {
      doc.rect(xPos, yPos, w, h).strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.rect(xPos, yPos, w, 12).fill('#eef2f7');
      doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY)
         .text(label.toUpperCase(), xPos + 6, yPos + 3, { width: w - 12 });
      doc.font('Helvetica').fontSize(10).fillColor(DARK)
         .text(value || '—', xPos + 6, yPos + 15, { width: w - 12, height: h - 16, ellipsis: true });
      return yPos + h;
    };

    // ── SECCIÓN 1: DATOS DEL TRABAJO ─────────────────────────
    y = sectionHeader('DATOS DEL TRABAJO', '◉', y) + 4;

    const col1 = M, col2 = M + CW / 3, col3 = M + (CW * 2 / 3);
    const colW = CW / 3;
    fieldRow('Número de Pedido', pedido, col1, y, colW, 36);
    fieldRow('Tipo de Trabajo', tipoLabel, col2, y, colW * 2, 36);
    y += 40;

    fieldRow('Fecha y Hora', fechaHora, col1, y, colW * 2, 30);
    fieldRow('Operario', operario || '—', col3, y, colW, 30);
    y += 34;

    // ── SECCIÓN 2: DATOS DEL CLIENTE ─────────────────────────
    y += 8;
    y = sectionHeader('DATOS DEL CLIENTE', '▣', y) + 4;
    fieldRow('Nombre Completo', clienteNombre || '—', col1, y, colW * 2, 30);
    fieldRow('DNI / NIF', clienteDni || '—', col3, y, colW, 30);
    y += 34;

    // ── SECCIÓN 3: OBSERVACIONES ─────────────────────────────
    y += 8;
    y = sectionHeader('OBSERVACIONES Y DETALLES DEL TRABAJO', '✎', y) + 4;
    const obsText = observaciones?.trim() || 'Sin observaciones registradas.';
    const obsLines = Math.max(3, Math.ceil(obsText.length / 90));
    const obsH = Math.min(obsLines * 14 + 16, 100);
    doc.rect(M, y, CW, obsH).strokeColor(BORDER).lineWidth(0.5).stroke();
    doc.rect(M, y, CW, 12).fill('#eef2f7');
    doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('OBSERVACIONES', M + 6, y + 3);
    doc.font('Helvetica').fontSize(10).fillColor(DARK)
       .text(obsText, M + 6, y + 16, { width: CW - 12, height: obsH - 20, lineGap: 2 });
    y += obsH + 4;

    // ── SECCIÓN 4: FIRMA DEL CLIENTE ─────────────────────────
    y += 8;
    y = sectionHeader('CONFORMIDAD Y FIRMA DEL CLIENTE', '✔', y) + 4;

    // Firma box
    const sigW = 220, sigH = 110;
    doc.rect(M, y, sigW, sigH).strokeColor(BORDER).lineWidth(0.5).stroke();
    doc.rect(M, y, sigW, 14).fill('#eef2f7');
    doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('FIRMA DEL CLIENTE', M + 6, y + 4);

    if (firma?.startsWith('data:image')) {
      try {
        const imgBuf = Buffer.from(firma.replace(/^data:image\/\w+;base64,/, ''), 'base64');
        doc.image(imgBuf, M + 4, y + 18, { width: sigW - 8, height: sigH - 22, fit: [sigW - 8, sigH - 22] });
      } catch (_) {
        doc.font('Helvetica').fontSize(9).fillColor('#aab').text('Firma no disponible', M + 10, y + 50);
      }
    } else {
      doc.font('Helvetica').fontSize(9).fillColor('#ccc').text('Sin firma registrada', M + 10, y + 55);
    }

    // Info declaratoria junto a firma
    const declX = M + sigW + 12;
    const declW = CW - sigW - 12;
    doc.rect(declX, y, declW, sigH).strokeColor(BORDER).lineWidth(0.5).stroke();
    doc.rect(declX, y, declW, 14).fill('#eef2f7');
    doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('DECLARACIÓN DE CONFORMIDAD', declX + 6, y + 4, { width: declW - 12 });
    doc.font('Helvetica').fontSize(8.5).fillColor(DARK)
       .text('El cliente abajo firmante declara haber recibido y revisado el trabajo descrito en el presente albarán, dando su conformidad con los trabajos realizados.', declX + 6, y + 20, { width: declW - 12, lineGap: 2 });
    doc.font('Helvetica').fontSize(8).fillColor(TEXTGRAY)
       .text(`Nombre: ${clienteNombre || '—'}`, declX + 6, y + 62, { width: declW - 12 });
    doc.font('Helvetica').fontSize(8).fillColor(TEXTGRAY)
       .text(`DNI/NIF: ${clienteDni || '—'}`, declX + 6, y + 75, { width: declW - 12 });
    doc.font('Helvetica').fontSize(8).fillColor(TEXTGRAY)
       .text(`Fecha: ${fechaHora}`, declX + 6, y + 88, { width: declW - 12 });
    y += sigH + 4;

    // ── SECCIÓN 5: GEOLOCALIZACIÓN (si disponible) ───────────
    if (geoData?.lat) {
      y += 8;
      y = sectionHeader('DATOS DE GEOLOCALIZACIÓN', '⊕', y) + 4;
      fieldRow('Latitud', String(geoData.lat), col1, y, colW, 28);
      fieldRow('Longitud', String(geoData.lng), col2, y, colW, 28);
      fieldRow('Precisión GPS', `±${geoData.accuracy} m`, col3, y, colW, 28);
      y += 32;
      const mapsUrl = `https://maps.google.com/?q=${geoData.lat},${geoData.lng}`;
      doc.rect(M, y, CW, 22).fill('#eef5fb').strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.font('Helvetica').fontSize(8).fillColor(BLUE)
         .text('Ver ubicación en Google Maps: ', M + 8, y + 7, { continued: true })
         .text(mapsUrl, { link: mapsUrl, underline: true });
      y += 26;
    }

    // ── FOOTER ───────────────────────────────────────────────
    const footerY = H - 38;
    doc.rect(0, footerY, W, 38).fill(DARK);
    doc.font('Helvetica').fontSize(7.5).fillColor('#7a9abf')
       .text(`ROURA & CEVASA  ·  ${tipoLabel}  ·  Pedido: ${pedido}  ·  Operario: ${operario || '—'}  ·  ${fechaHora}`, M, footerY + 8, { width: CW, align: 'center' });
    doc.font('Helvetica').fontSize(6.5).fillColor('#4a6478')
       .text('Documento generado automáticamente por el sistema FTP Manager · Uso interno exclusivo', M, footerY + 22, { width: CW, align: 'center' });

    // Page number
    doc.font('Helvetica').fontSize(7).fillColor('#6a8aaa').text('Pág. 1', W - M - 30, footerY + 14);

    doc.end();
  });
}


// ── BUSCAR PEDIDO EN EXCEL ────────────────────────────────────
app.get('/api/buscar-pedido', requireAuth, async (req, res) => {
  const numeroPedido = (req.query.pedido || '').trim();
  if (!numeroPedido) return res.status(400).json({ success: false, error: 'Falta número de pedido' });

  const EXCEL_FILE = 'vw_segplazo_052025.xlsx';
  const REMOTE_PATH = path.posix.join(BASE_PATH, EXCEL_FILE);

  const client = new ftp.Client(60000);
  client.ftp.verbose = false;

  try {
    await client.access(FTP_CONFIG);

    // Download the Excel file into a buffer via PassThrough stream
    const { PassThrough } = require('stream');
    const pt = new PassThrough();
    const chunks = [];
    pt.on('data', c => chunks.push(c));
    const finished = new Promise((resolve, reject) => { pt.on('end', resolve); pt.on('error', reject); });
    await client.downloadTo(pt, REMOTE_PATH);
    await finished;
    const buffer = Buffer.concat(chunks);

    // Parse with xlsx
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheet = workbook.Sheets['DatosX3'];
    if (!sheet) return res.status(404).json({ success: false, error: 'No se encontro la hoja DatosX3 en el archivo Excel' });

    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

    // Search column A (index 0) for the order number, skip header row
    let found = null;
    for (let i = 1; i < rows.length; i++) {
      const cellA = String(rows[i][0] || '').trim();
      if (cellA === numeroPedido) {
        const row = rows[i];
        found = {
          SIE:        String(row[1]  || '').trim(),  // Columna B
          RefCli:     String(row[2]  || '').trim(),  // Columna C
          PA:         String(row[4]  || '').trim(),  // Columna E
          Ciudad:     String(row[5]  || '').trim(),  // Columna F
          Direccion:  String(row[6]  || '').trim(),  // Columna G
          Provincia:  String(row[7]  || '').trim(),  // Columna H
          Descripcion: String(row[8] || '').trim(),  // Columna I
        };
        break;
      }
    }

    if (!found) {
      return res.json({ success: false, found: false, error: 'El pedido "' + numeroPedido + '" no existe en la tabla DatosX3' });
    }

    res.json({ success: true, found: true, datos: found });
  } catch (e) {
    res.status(500).json({ success: false, error: 'Error accediendo al Excel: ' + e.message });
  } finally {
    client.close();
  }
});



// ── HELPER: descargar archivo FTP a Buffer ────────────────────
async function ftpDownloadBuffer(client, remotePath) {
  const pt = new PassThrough();
  const chunks = [];
  pt.on('data', c => chunks.push(c));
  const done = new Promise((res, rej) => { pt.on('end', res); pt.on('error', rej); });
  await client.downloadTo(pt, remotePath);
  await done;
  return Buffer.concat(chunks);
}

// ── HELPER: listar archivos de una carpeta FTP (sin error si no existe) ──
async function ftpListSafe(client, remotePath) {
  try {
    await client.cd(remotePath);
    return await client.list();
  } catch (_) { return []; }
}

// ── HELPER: es imagen por extensión ──────────────────────────
function esImagen(nombre) {
  return /\.(jpe?g|png|gif|webp|bmp)$/i.test(nombre);
}

// ── CERTIFICACION ─────────────────────────────────────────────
app.post('/api/certificacion', requireAuth, async (req, res) => {
  const { pedido, datosExcel, albaranPath, timestamp, operario: opReq } = req.body;
  if (!pedido) return res.status(400).json({ success: false, error: 'Falta número de pedido' });

  const operario = opReq || req.user.name;
  const fechaHora = timestamp || new Date().toLocaleString('es-ES');

  // Rutas FTP (nuevas rutas centralizadas)
  const dirAntes       = path.posix.join(BASE_PATH, 'Fotos Inicio');
  const dirFinal       = path.posix.join(BASE_PATH, 'Fotos Fin');
  const dirAlbaran     = path.posix.join(BASE_PATH, 'Albaranes Final Trabajo');
  const dirCert        = path.posix.join(BASE_PATH, 'Certificacion Final Trabajo');

  const client = new ftp.Client(300000);
  client.ftp.verbose = false;

  try {
    await client.access(FTP_CONFIG);

    // ── Descargar fotos antes ──────────────────────────────
    const listAntes = await ftpListSafe(client, dirAntes);
    const fotosAntesBufs = [];
    for (const f of listAntes.filter(f => f.type !== 2 && esImagen(f.name))) {
      try {
        const buf = await ftpDownloadBuffer(client, path.posix.join(dirAntes, f.name));
        fotosAntesBufs.push({ name: f.name, buf });
      } catch (_) {}
    }

    // ── Descargar fotos final ──────────────────────────────
    const listFinal = await ftpListSafe(client, dirFinal);
    const fotosFinalBufs = [];
    for (const f of listFinal.filter(f => f.type !== 2 && esImagen(f.name))) {
      try {
        const buf = await ftpDownloadBuffer(client, path.posix.join(dirFinal, f.name));
        fotosFinalBufs.push({ name: f.name, buf });
      } catch (_) {}
    }

    // ── Descargar el albarán más reciente de albaran_instalacion ──
    let albaranBuf = null;
    const listAlb = await ftpListSafe(client, dirAlbaran);
    const pdfsAlb = listAlb
      .filter(f => f.type !== 2 && /\.pdf$/i.test(f.name))
      .sort((a, b) => new Date(b.modifiedAt) - new Date(a.modifiedAt));
    if (pdfsAlb.length > 0) {
      try {
        albaranBuf = await ftpDownloadBuffer(client, path.posix.join(dirAlbaran, pdfsAlb[0].name));
      } catch (_) {}
    }

    // ── Generar PDF de Certificación ──────────────────────
    const pdfBuf = await generarCertificacion({
      pedido, datosExcel, fechaHora, operario,
      fotosAntesBufs, fotosFinalBufs, albaranBuf
    });

    // ── Subir al FTP en carpeta Estándar como ZIP ─────────
    await ensureDir(client, dirCert);
    const tsCert  = tsNombre(req.user.id || '');
    const pdfName = `${String(pedido).trim()}_${tsCert}_Certificacion_Final_Trabajo.pdf`;
    await client.uploadFrom(Readable.from(pdfBuf), path.posix.join(dirCert, pdfName));

    res.json({ success: true, fileName: pdfName, ruta: `${dirCert}/${pdfName}` });
  } catch (e) {
    console.error('Error certificacion:', e);
    res.status(500).json({ success: false, error: 'Error generando certificación: ' + e.message });
  } finally {
    client.close();
  }
});

// ── HELPER: convertir PDF a imágenes PNG con pdftoppm ────────
async function pdfToImages(pdfBuffer) {
  const tmpDir    = fs.mkdtempSync(path.join(os.tmpdir(), 'albaran_'));
  const tmpPdf    = path.join(tmpDir, 'albaran.pdf');
  const outPrefix = path.join(tmpDir, 'page');
  const images    = [];
  try {
    fs.writeFileSync(tmpPdf, pdfBuffer);
    // -r 150 = 150 dpi  (buen equilibrio calidad/tamaño en PDF final)
    execSync(`pdftoppm -r 150 -png "${tmpPdf}" "${outPrefix}"`, { timeout: 60000 });
    const files = fs.readdirSync(tmpDir)
      .filter(f => f.startsWith('page') && f.endsWith('.png'))
      .sort();
    for (const f of files) {
      images.push(fs.readFileSync(path.join(tmpDir, f)));
    }
  } finally {
    try { fs.rmSync(tmpDir, { recursive: true, force: true }); } catch (_) {}
  }
  return images;
}

// ── GENERADOR PDF CERTIFICACION ───────────────────────────────
async function generarCertificacion({ pedido, datosExcel, fechaHora, operario, fotosAntesBufs, fotosFinalBufs, albaranBuf }) {

  // Convertir albarán PDF → imágenes antes de crear el documento PDFKit
  let albaranImagenes = [];
  if (albaranBuf) {
    try { albaranImagenes = await pdfToImages(albaranBuf); } catch (e) { console.warn('pdf2img error:', e.message); }
  }

  return new Promise((resolve, reject) => {
    const doc = new PDFDocument({ margin: 0, size: 'A4', autoFirstPage: true });
    const bufs = [];
    doc.on('data', c => bufs.push(c));
    doc.on('end',  () => resolve(Buffer.concat(bufs)));
    doc.on('error', reject);

    // ── Constantes de página ───────────────────────────────
    const W  = 595;   // A4 puntos
    const H  = 842;
    const M  = 30;    // margen lateral
    const CW = W - M * 2;

    // Zona segura para fotos (entre cabecera y pie)
    const HEADER_H = 90;   // altura cabecera en páginas de foto
    const FOOTER_H = 28;   // altura pie de página
    const FOTO_AREA_Y = HEADER_H + 8;                    // inicio zona foto
    const FOTO_AREA_H = H - FOTO_AREA_Y - FOOTER_H - 8; // alto disponible foto

    // ── Colores ────────────────────────────────────────────
    const BLUE     = '#1B6CA8';
    const DARK     = '#1a2332';
    const BORDER   = '#d0d8e4';
    const TEXTGRAY = '#5a6478';
    const WHITE    = '#ffffff';

    // ── Logo ───────────────────────────────────────────────
    let logoBuf = null;
    if (LOGO_B64) {
      try { logoBuf = Buffer.from(LOGO_B64.replace(/^data:image\/\w+;base64,/, ''), 'base64'); } catch (_) {}
    }

    // ══════════════════════════════════════════════════════
    // HELPERS
    // ══════════════════════════════════════════════════════

    /** Cabecera compacta para páginas de foto/albarán */
    const drawMiniHeader = (secLabel) => {
      doc.rect(0, 0, W, HEADER_H).fill(DARK);
      if (logoBuf) { try { doc.image(logoBuf, M, 10, { height: 68, fit: [68, 68] }); } catch (_) {} }
      doc.font('Helvetica-Bold').fontSize(14).fillColor(WHITE)
         .text('ROURA & CEVASA', M + 80, 14, { width: 240 });
      doc.font('Helvetica').fontSize(8).fillColor('#a8c4e0')
         .text('Gestión Técnica y Documentación', M + 80, 34, { width: 240 });
      doc.font('Helvetica').fontSize(7.5).fillColor('#7a9abf')
         .text(fechaHora, M + 80, 50, { width: 240 });
      // Badge sección derecha
      doc.roundedRect(W - M - 130, 12, 130, 64, 6).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(9).fillColor(WHITE)
         .text(secLabel.toUpperCase(), W - M - 130, 30, { width: 130, align: 'center' });
      doc.font('Helvetica').fontSize(7).fillColor('#cde4f5')
         .text('Pedido: ' + pedido, W - M - 130, 50, { width: 130, align: 'center' });
      // Franja azul inferior
      doc.rect(0, HEADER_H, W, 4).fill(BLUE);
    };

    /** Pie de página con número de foto y sección */
    const drawFotoPie = (num, total, secLabel) => {
      const py = H - FOOTER_H;
      doc.rect(0, py, W, FOOTER_H).fill(DARK);
      doc.font('Helvetica').fontSize(7.5).fillColor('#7a9abf')
         .text(
           `${secLabel}  ·  Foto ${num} de ${total}  ·  Pedido: ${pedido}  ·  ${fechaHora}`,
           M, py + 8, { width: CW, align: 'center' }
         );
    };

    /** Pie genérico */
    const drawPie = (texto) => {
      const py = H - FOOTER_H;
      doc.rect(0, py, W, FOOTER_H).fill(DARK);
      doc.font('Helvetica').fontSize(7.5).fillColor('#7a9abf')
         .text(texto, M, py + 8, { width: CW, align: 'center' });
    };

    /** Banner de sección a página completa – primera página de cada sección */
    const drawSectionCover = (num, titulo, subtitulo) => {
      // Fondo degradado oscuro toda la página
      doc.rect(0, 0, W, H).fill(DARK);
      // Franja azul superior
      doc.rect(0, 0, W, 8).fill(BLUE);
      // Número grande
      doc.font('Helvetica-Bold').fontSize(120).fillColor(BLUE)
         .text(String(num), 0, H / 2 - 130, { width: W, align: 'center', opacity: 0.25 });
      // Logo centrado
      if (logoBuf) {
        try { doc.image(logoBuf, W / 2 - 50, H / 2 - 120, { width: 100, height: 100, fit: [100, 100], align: 'center', valign: 'center' }); } catch (_) {}
      }
      // Línea separadora
      doc.rect(W / 2 - 80, H / 2 - 10, 160, 3).fill(BLUE);
      // Título
      doc.font('Helvetica-Bold').fontSize(28).fillColor(WHITE)
         .text(titulo.toUpperCase(), 0, H / 2 + 4, { width: W, align: 'center' });
      // Subtítulo
      doc.font('Helvetica').fontSize(11).fillColor('#a8c4e0')
         .text(subtitulo, 0, H / 2 + 46, { width: W, align: 'center' });
      // Pedido abajo
      doc.font('Helvetica').fontSize(9).fillColor('#7a9abf')
         .text('Pedido nº ' + pedido + '  ·  ' + fechaHora, 0, H - 50, { width: W, align: 'center' });
      // Franja azul inferior
      doc.rect(0, H - 8, W, 8).fill(BLUE);
    };

    /** Una foto centrada al máximo en la zona disponible */
    const drawFotoPagina = (imgBuf, nombre, num, total, secLabel) => {
      drawMiniHeader(secLabel);
      drawFotoPie(num, total, secLabel);

      // Zona disponible para la foto
      const areaX = M;
      const areaY = FOTO_AREA_Y;
      const areaW = CW;
      const areaH = FOTO_AREA_H;

      // Fondo suave para la zona foto
      doc.rect(areaX, areaY, areaW, areaH).fill('#0d1929');

      // Número + nombre de archivo en una tira encima de la foto
      const tiraH = 20;
      doc.rect(areaX, areaY, areaW, tiraH).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(8).fillColor(WHITE)
         .text(`Foto ${num} / ${total} — ${nombre}`, areaX + 8, areaY + 5, { width: areaW - 16, ellipsis: true });

      // Imagen: máxima dentro del área, respetando aspecto
      const imgAreaY = areaY + tiraH + 6;
      const imgAreaH = areaH - tiraH - 12;
      const imgAreaW = areaW - 12;

      try {
        doc.image(imgBuf, areaX + 6, imgAreaY, {
          width:  imgAreaW,
          height: imgAreaH,
          fit:    [imgAreaW, imgAreaH],
          align:  'center',
          valign: 'center'
        });
      } catch (_) {
        doc.font('Helvetica').fontSize(12).fillColor('#555')
           .text('Imagen no disponible', areaX, imgAreaY + imgAreaH / 2 - 8, { width: areaW, align: 'center' });
      }
    };

    /** Campo de datos */
    const fieldRow = (label, value, x, y, w, h = 36) => {
      doc.rect(x, y, w, h).strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.rect(x, y, w, 13).fill('#eef2f7');
      doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY)
         .text(label.toUpperCase(), x + 6, y + 3, { width: w - 12 });
      doc.font('Helvetica').fontSize(10).fillColor(DARK)
         .text(value || '—', x + 6, y + 16, { width: w - 12, height: h - 18, ellipsis: true });
    };

    // ══════════════════════════════════════════════════════
    // PÁGINA 1 — DATOS DE INSTALACIÓN
    // ══════════════════════════════════════════════════════
    {
      // Cabecera completa primera página
      doc.rect(0, 0, W, 110).fill(DARK);
      if (logoBuf) { try { doc.image(logoBuf, M, 18, { height: 74, fit: [74, 74] }); } catch (_) {} }
      doc.font('Helvetica-Bold').fontSize(18).fillColor(WHITE)
         .text('ROURA & CEVASA', M + 88, 22, { width: 260 });
      doc.font('Helvetica').fontSize(9.5).fillColor('#a8c4e0')
         .text('Gestión Técnica y Documentación', M + 88, 46, { width: 260 });
      doc.font('Helvetica').fontSize(8).fillColor('#7a9abf')
         .text('Imagen Corporativa · Certificación de Instalación', M + 88, 64, { width: 260 });
      doc.roundedRect(W - M - 148, 20, 148, 70, 6).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(7).fillColor('#cde4f5')
         .text('CERTIFICACIÓN', W - M - 148, 30, { width: 148, align: 'center' });
      doc.font('Helvetica-Bold').fontSize(12).fillColor(WHITE)
         .text('INSTALACIÓN', W - M - 148, 46, { width: 148, align: 'center' });
      doc.font('Helvetica').fontSize(7.5).fillColor('#cde4f5')
         .text(fechaHora, W - M - 148, 68, { width: 148, align: 'center' });
      doc.rect(0, 110, W, 5).fill(BLUE);

      doc.y = 128;

      // Banner sección 1
      doc.rect(M, doc.y, CW, 28).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(12).fillColor(WHITE)
         .text('1 — INSTALACIÓN', M + 12, doc.y + 8, { width: CW - 24 });
      doc.y += 36;

      // Sub-cabecera obra
      doc.rect(M, doc.y, CW, 28).fill('#e8f2fb');
      doc.rect(M, doc.y, 4, 28).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(11).fillColor(DARK)
         .text('Instalación de Imagen Corporativa', M + 14, doc.y + 8, { width: CW - 20 });
      doc.y += 36;

      // Grid de campos
      const dx   = datosExcel || {};
      const half = CW / 2;
      let   gy   = doc.y;

      fieldRow('Número de Pedido',    String(pedido),      M,         gy, half,    36);
      fieldRow('SIE',                 dx.SIE  || '—',      M + half,  gy, half,    36);
      gy += 40;
      fieldRow('Referencia Cliente',  dx.RefCli || '—',    M,         gy, half,    36);
      fieldRow('PA',                  dx.PA || '—',        M + half,  gy, half,    36);
      gy += 40;
      fieldRow('Ciudad',              dx.Ciudad || '—',    M,         gy, CW*0.42, 36);
      fieldRow('Provincia',           dx.Provincia || '—', M+CW*0.42, gy, CW*0.33, 36);
      fieldRow('Fecha',               fechaHora,           M+CW*0.75, gy, CW*0.25, 36);
      gy += 40;
      fieldRow('Dirección',           dx.Direccion || '—', M,         gy, CW,      36);
      gy += 40;
      fieldRow('Descripción',         dx.Descripcion || '—', M,       gy, CW,      36);
      gy += 40;
      fieldRow('Operario',            operario || '—',     M,         gy, CW,      32);
      gy += 36;

      // Separador
      doc.rect(M, gy + 10, CW, 1).fill(BORDER);
      gy += 20;

      // Párrafo corporativo
      doc.font('Helvetica-Bold').fontSize(10).fillColor(BLUE)
         .text('ROURA & CEVASA — Imagen Corporativa', M, gy, { width: CW, align: 'center' });
      gy += 18;
      doc.font('Helvetica').fontSize(8.5).fillColor(TEXTGRAY)
         .text(
           'Este documento certifica la correcta ejecución e instalación de los elementos de imagen ' +
           'corporativa en las instalaciones indicadas, conforme a los estándares de calidad y ' +
           'normativa vigente de ROURA & CEVASA.',
           M, gy, { width: CW, align: 'center', lineGap: 2 }
         );
      gy += 44;

      // Badge estado
      doc.rect(M, gy, CW, 40).fill('#e8f5ef').strokeColor('#0ea569').lineWidth(1).stroke();
      doc.circle(M + 22, gy + 20, 9).fill('#0ea569');
      doc.font('Helvetica-Bold').fontSize(11).fillColor('#065f46')
         .text('INSTALACIÓN COMPLETADA Y DOCUMENTADA', M + 40, gy + 13, { width: CW - 50 });

      // Pie página 1
      drawPie(`1 — Instalación  ·  Pedido: ${pedido}  ·  ${fechaHora}`);
    }

    // ══════════════════════════════════════════════════════
    // SECCIÓN 2 — FOTOS ANTES
    // ══════════════════════════════════════════════════════
    {
      // Portada sección
      doc.addPage();
      drawSectionCover(2, 'Fotos Antes', `${fotosAntesBufs.length} fotografía(s) del estado previo a la instalación`);

      // Una foto por página
      if (fotosAntesBufs.length === 0) {
        doc.addPage();
        drawMiniHeader('2 — Fotos Antes');
        doc.font('Helvetica').fontSize(13).fillColor(TEXTGRAY)
           .text('No se encontraron fotografías de "antes" para este pedido.', M, H / 2 - 20, { width: CW, align: 'center' });
        drawPie(`2 — Fotos Antes  ·  Pedido: ${pedido}  ·  ${fechaHora}`);
      } else {
        fotosAntesBufs.forEach((foto, i) => {
          doc.addPage();
          drawFotoPagina(foto.buf, foto.name, i + 1, fotosAntesBufs.length, '2 — Fotos Antes');
        });
      }
    }

    // ══════════════════════════════════════════════════════
    // SECCIÓN 3 — ALBARÁN TRABAJOS
    // ══════════════════════════════════════════════════════
    {
      // Portada sección
      doc.addPage();
      drawSectionCover(3, 'Albarán Trabajos', 'Albarán de instalación firmado por el cliente');

      if (albaranImagenes.length === 0) {
        doc.addPage();
        drawMiniHeader('3 — Albarán Trabajos');
        doc.font('Helvetica').fontSize(13).fillColor(TEXTGRAY)
           .text('No se encontró el albarán para este pedido.', M, H / 2 - 20, { width: CW, align: 'center' });
        drawPie(`3 — Albarán Trabajos  ·  Pedido: ${pedido}  ·  ${fechaHora}`);
      } else {
        // Una página del albarán por página del PDF, máxima calidad
        albaranImagenes.forEach((imgBuf, i) => {
          doc.addPage();
          drawMiniHeader('3 — Albarán Trabajos');
          drawFotoPie(i + 1, albaranImagenes.length, '3 — Albarán Trabajos');

          // Página del albarán centrada al máximo
          const areaY = FOTO_AREA_Y;
          const areaH = FOTO_AREA_H - 4;
          const areaW = CW;
          doc.rect(M, areaY, areaW, areaH).fill('#f8fafc');
          try {
            doc.image(imgBuf, M + 2, areaY + 2, {
              width:  areaW - 4,
              height: areaH - 4,
              fit:    [areaW - 4, areaH - 4],
              align:  'center',
              valign: 'center'
            });
          } catch (_) {
            doc.font('Helvetica').fontSize(11).fillColor(TEXTGRAY)
               .text('Página de albarán no disponible', M, areaY + areaH / 2 - 8, { width: CW, align: 'center' });
          }
        });
      }
    }

    // ══════════════════════════════════════════════════════
    // SECCIÓN 4 — FOTOS FINAL
    // ══════════════════════════════════════════════════════
    {
      // Portada sección
      doc.addPage();
      drawSectionCover(4, 'Fotos Final', `${fotosFinalBufs.length} fotografía(s) del resultado de la instalación`);

      if (fotosFinalBufs.length === 0) {
        doc.addPage();
        drawMiniHeader('4 — Fotos Final');
        doc.font('Helvetica').fontSize(13).fillColor(TEXTGRAY)
           .text('No se encontraron fotografías de "final" para este pedido.', M, H / 2 - 20, { width: CW, align: 'center' });
        drawPie(`4 — Fotos Final  ·  Pedido: ${pedido}  ·  ${fechaHora}`);
      } else {
        fotosFinalBufs.forEach((foto, i) => {
          doc.addPage();
          drawFotoPagina(foto.buf, foto.name, i + 1, fotosFinalBufs.length, '4 — Fotos Final');
        });
      }
    }

    doc.end();
  });
}



// ══════════════════════════════════════════════════════════════
// ── CFO — VISITA CFO ──────────────────────────────────────────
// ══════════════════════════════════════════════════════════════

/*
  POST /api/cfo
  Body (JSON, puede ser grande por las imágenes en base64):
  {
    pedido: "12345",
    datosExcel: { SIE, RefCli, PA, Ciudad, Direccion, Provincia, Descripcion },
    timestamp: "dd/mm/aaaa hh:mm:ss",
    incidencias: [
      {
        numero: 1,
        fotosDoc:       [ { name, dataUrl } ],   // campo 1
        descripcion:    "texto",                  // campo 2
        tipologia:      "material|equipo|documental",
        infAdicional:   [ { name, dataUrl } ],   // campo 4
        descAdicional:  "texto"                   // campo 5
      },
      ...
    ]
  }
  Genera:
    - PDF  → /www/DMA/Apertura Objeciones/{pedido}_cfo_{ts}.zip
    - JSON → /www/DMA/CierresCFO/{pedido}_cfo_{ts}.json   (sin comprimir)
*/

// ── CIERRE OBJECIONES CFO — listar JSONs pendientes en Apertura Objeciones json ──
app.get('/api/cfo-listar-pendientes', requireAuth, async (req, res) => {
  const { pedido } = req.query;
  if (!pedido) return res.status(400).json({ success: false, error: 'Falta número de pedido' });

  const client = new ftp.Client(60000);
  client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);
    const dirVisita = path.posix.join(BASE_PATH, 'Apertura Objeciones json');
    await ensureDir(client, dirVisita);
    const lista = await client.list(dirVisita);
    const prefix = String(pedido).trim().toLowerCase();
    const archivos = lista
      .filter(f => f.type === 1 && f.name.toLowerCase().startsWith(prefix) && f.name.toLowerCase().endsWith('.json'))
      .map(f => f.name);
    res.json({ success: true, archivos });
  } catch(e) {
    res.status(500).json({ success: false, error: 'Error listando Apertura Objeciones json: ' + e.message });
  } finally {
    client.close();
  }
});

// ── CIERRE OBJECIONES CFO — cargar JSON concreto de Apertura Objeciones json ─
app.get('/api/cfo-cargar-json', requireAuth, async (req, res) => {
  const { nombre } = req.query;
  if (!nombre) return res.status(400).json({ success: false, error: 'Falta nombre de archivo' });

  const client = new ftp.Client(120000);
  client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);
    const remotePath = path.posix.join(BASE_PATH, 'Apertura Objeciones json', nombre);
    const chunks = [];
    const writable = new (require('stream').Writable)({
      write(chunk, _enc, cb) { chunks.push(chunk); cb(); }
    });
    await client.downloadTo(writable, remotePath);
    const jsonStr = Buffer.concat(chunks).toString('utf8');
    const jsonData = JSON.parse(jsonStr);
    res.json({ success: true, json: jsonData });
  } catch(e) {
    res.status(500).json({ success: false, error: 'Error cargando JSON: ' + e.message });
  } finally {
    client.close();
  }
});

// ── CIERRE OBJECIONES CFO — generar PDF+JSON, mover archivo ────────────────
app.post('/api/cierre-cfo', requireAuth, async (req, res) => {
  const { pedido, datosExcel, fechaHoraVisita, timestamp, resoluciones, jsonFileName } = req.body;
  if (!pedido) return res.status(400).json({ success: false, error: 'Falta número de pedido' });
  if (!Array.isArray(resoluciones) || resoluciones.length === 0)
    return res.status(400).json({ success: false, error: 'No hay resoluciones' });
  if (!jsonFileName) return res.status(400).json({ success: false, error: 'Falta nombre de archivo JSON' });

  const operario  = req.user.name;
  const fechaHora = timestamp || new Date().toLocaleString('es-ES');
  const ts        = Date.now();

  // 1. Generar PDF
  let pdfBuf;
  try {
    pdfBuf = await generarCierreCFOPDF({ pedido, datosExcel, resoluciones, fechaHora, fechaHoraVisita, operario });
  } catch(e) {
    return res.status(500).json({ success: false, error: 'Error generando PDF cierre: ' + e.message });
  }

  // 2. Generar JSON (CON imágenes base64 para permitir modificaciones futuras)
  const jsonData = {
    pedido, datosExcel, fechaHoraVisita, fechaHoraCierre: fechaHora, operario,
    generadoEn: new Date().toISOString(),
    version: 1,
    totalResueltas: resoluciones.length,
    resoluciones: resoluciones.map(r => ({
      incidenciaNumero: r.incidenciaNumero,
      tipologia:        r.incidencia?.tipologia    || '',
      descripcion:      r.incidencia?.descripcion  || '',
      descAdicional:    r.incidencia?.descAdicional || '',
      anotacion:        r.anotacion                || '',
      fechaResolucion:  r.fechaResolucion          || '',
      documentos:       (r.documentacion || []).map(f => f.name),
      documentacion:    r.documentacion            || [],  // incluye dataUrl
      incidencia:       r.incidencia               || {}   // incluye fotosDoc/infAdicional con dataUrl
    }))
  };
  const jsonBuf = Buffer.from(JSON.stringify(jsonData, null, 2), 'utf8');

  // 3. Subir al FTP y mover archivo visita
  const client = new ftp.Client(300000);
  client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);

    // PDF → /www/DMA/Cierre Objeciones
    const dirCierre  = path.posix.join(BASE_PATH, 'Cierre Objeciones');
    await ensureDir(client, dirCierre);
    const tsCierre   = tsNombre(req.user.id || '');
    const pdfName    = `${String(pedido).trim()}_${tsCierre}_Informe_Cierre_Objeciones_CFO.pdf`;
    await client.uploadFrom(Readable.from(pdfBuf), path.posix.join(dirCierre, pdfName));

    // JSON → /www/DMA/Cierre Objeciones json
    const dirCierreJSON = path.posix.join(BASE_PATH, 'Cierre Objeciones json');
    await ensureDir(client, dirCierreJSON);
    const jsonName   = `${String(pedido).trim()} Informe Cierre Objeciones CFO_${tsCierre}.json`;
    await client.uploadFrom(Readable.from(jsonBuf), path.posix.join(dirCierreJSON, jsonName));

    // Mover JSON Apertura Objeciones json → Historico Visitas CFO
    const dirHistorico = path.posix.join(BASE_PATH, 'Historico Visitas CFO');
    await ensureDir(client, dirHistorico);
    const srcPath = path.posix.join(BASE_PATH, 'Apertura Objeciones json', jsonFileName);
    const dstPath = path.posix.join(dirHistorico, jsonFileName);
    const chunksMov = [];
    const wMov = new (require('stream').Writable)({
      write(chunk, _enc, cb) { chunksMov.push(chunk); cb(); }
    });
    await client.downloadTo(wMov, srcPath);
    const fileBuf = Buffer.concat(chunksMov);
    await client.uploadFrom(Readable.from(fileBuf), dstPath);
    await client.remove(srcPath);

    res.json({ success: true, pdfFile: pdfName, jsonFile: jsonName,
      rutaPDF:  `${dirCierre}/${pdfName}`,
      rutaJSON: `${dirCierreJSON}/${jsonName}` });
  } catch(e) {
    res.status(500).json({ success: false, error: 'Error en FTP cierre CFO: ' + e.message });
  } finally {
    client.close();
  }
});


// ── GENERADOR PDF CIERRE CFO ──────────────────────────────────────────────
async function generarCierreCFOPDF({ pedido, datosExcel, resoluciones, fechaHora, fechaHoraVisita, operario }) {
  return new Promise((resolve, reject) => {
    const doc  = new PDFDocument({ margin: 0, size: 'A4', autoFirstPage: true });
    const bufs = [];
    doc.on('data',  c => bufs.push(c));
    doc.on('end',   () => resolve(Buffer.concat(bufs)));
    doc.on('error', reject);

    const W  = 595, H = 842, M = 32;
    const CW = W - M * 2;
    const BLUE     = '#1B6CA8';
    const DARK     = '#1a2332';
    const BORDER   = '#d0d8e4';
    const TEXTGRAY = '#5a6478';
    const WHITE    = '#ffffff';
    const LGRAY    = '#f4f7fb';
    const GREEN    = '#1e8a4c';

    let logoBuf = null;
    if (LOGO_B64) {
      try { logoBuf = Buffer.from(LOGO_B64.replace(/^data:image\/\w+;base64,/, ''), 'base64'); } catch (_) {}
    }
    const dx = datosExcel || {};

    const drawPageHeader = (label) => {
      doc.rect(0, 0, W, 90).fill(DARK);
      if (logoBuf) { try { doc.image(logoBuf, M, 10, { height: 68, fit: [68, 68] }); } catch (_) {} }
      doc.font('Helvetica-Bold').fontSize(15).fillColor(WHITE).text('ROURA & CEVASA', M + 78, 14, { width: 220 });
      doc.font('Helvetica').fontSize(8).fillColor('#a8c4e0').text('Cierre Objeciones CFO — Informe de Resolución', M + 78, 34, { width: 220 });
      doc.font('Helvetica').fontSize(7.5).fillColor('#7a9abf').text(fechaHora, M + 78, 50, { width: 220 });
      doc.roundedRect(W - M - 132, 10, 132, 68, 6).fill(GREEN);
      doc.font('Helvetica-Bold').fontSize(9).fillColor(WHITE).text(label.toUpperCase(), W - M - 132, 28, { width: 132, align: 'center' });
      doc.font('Helvetica').fontSize(7).fillColor('#cde4f5').text('Pedido: ' + pedido, W - M - 132, 50, { width: 132, align: 'center' });
      doc.rect(0, 90, W, 4).fill(GREEN);
    };

    const drawFooter = (pageLabel) => {
      const py = H - 26;
      doc.rect(0, py, W, 26).fill(DARK);
      doc.font('Helvetica').fontSize(7).fillColor('#7a9abf')
         .text(`${pageLabel}  ·  Pedido: ${pedido}  ·  Operario: ${operario}  ·  ${fechaHora}`, M, py + 8, { width: CW, align: 'center' });
    };

    const fieldBox = (label, value, x, y, w, h = 34) => {
      doc.rect(x, y, w, h).strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.rect(x, y, w, 12).fill(LGRAY);
      doc.font('Helvetica-Bold').fontSize(6.5).fillColor(TEXTGRAY).text(label.toUpperCase(), x + 5, y + 3, { width: w - 10 });
      doc.font('Helvetica').fontSize(9).fillColor(DARK).text(String(value || '—'), x + 5, y + 15, { width: w - 10, height: h - 17, ellipsis: true });
    };

    // ══ PÁGINA 1 — PORTADA ══
    drawPageHeader('CIERRE CFO');
    let y = 106;

    doc.rect(M, y, CW, 30).fill(GREEN);
    doc.font('Helvetica-Bold').fontSize(12).fillColor(WHITE)
       .text('CIERRE OBJECIONES CFO — INFORME DE INCIDENCIAS RESUELTAS', M + 10, y + 9, { width: CW - 20 });
    y += 38;

    doc.rect(M, y, CW, 26).fill(LGRAY); doc.rect(M, y, 4, 26).fill(GREEN);
    doc.font('Helvetica-Bold').fontSize(10).fillColor(DARK).text('Datos de la Instalación', M + 12, y + 8, { width: CW - 20 });
    y += 32;

    const half = CW / 2;
    fieldBox('Número de Pedido', pedido,            M,        y, half,      34);
    fieldBox('SIE',              dx.SIE || '—',     M + half, y, half,      34); y += 38;
    fieldBox('Referencia Cliente', dx.RefCli || '—', M,       y, half,      34);
    fieldBox('PA',               dx.PA    || '—',   M + half, y, half,      34); y += 38;
    fieldBox('Ciudad',    dx.Ciudad    || '—', M,            y, CW*0.4,  34);
    fieldBox('Provincia', dx.Provincia || '—', M + CW*0.4,   y, CW*0.35, 34);
    fieldBox('Fecha Cierre', fechaHora,        M + CW*0.75,  y, CW*0.25, 34); y += 38;
    fieldBox('Dirección',   dx.Direccion   || '—', M, y, CW, 34); y += 38;
    fieldBox('Descripción', dx.Descripcion  || '—', M, y, CW, 34); y += 38;
    fieldBox('Operario',    operario || '—',        M, y, CW * 0.6, 30);
    fieldBox('Fecha Visita CFO', fechaHoraVisita || '—', M + CW * 0.6, y, CW * 0.4, 30); y += 36;

    doc.rect(M, y + 6, CW, 1).fill(BORDER); y += 16;

    // Tabla resumen
    doc.rect(M, y, CW, 26).fill(GREEN);
    doc.font('Helvetica-Bold').fontSize(10).fillColor(WHITE)
       .text(`Resumen: ${resoluciones.length} incidencia(s) resuelta(s)`, M + 10, y + 8, { width: CW - 20 });
    y += 32;

    const colW = [40, 190, 155, 80, 66];
    const cols = [M, M+colW[0], M+colW[0]+colW[1], M+colW[0]+colW[1]+colW[2], M+colW[0]+colW[1]+colW[2]+colW[3]];
    ['Nº','Descripción','Tipología','Resolución','Docs'].forEach((h, i) => {
      doc.rect(cols[i], y, colW[i], 20).fill('#e2f4ea').strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.font('Helvetica-Bold').fontSize(7.5).fillColor(DARK).text(h, cols[i]+4, y+6, { width: colW[i]-8 });
    });
    y += 20;

    const TIPMAP = { material:'Nec. Material', equipo:'Eq. resuelve', documental:'Documental' };
    resoluciones.slice(0, 35).forEach(r => {
      const rh = 18;
      if (y + rh > H - 40) { drawFooter('Portada Cierre CFO'); doc.addPage(); drawPageHeader('Cierre CFO — Resumen'); y = 106; }
      const rowBg = (r.incidenciaNumero % 2 === 0) ? '#f8fafc' : WHITE;
      [0,1,2,3,4].forEach(i => doc.rect(cols[i], y, colW[i], rh).fill(rowBg).strokeColor(BORDER).lineWidth(0.3).stroke());
      doc.font('Helvetica').fontSize(7.5).fillColor(DARK).text(String(r.incidenciaNumero), cols[0]+4, y+5, { width: colW[0]-8 });
      doc.font('Helvetica').fontSize(7).fillColor(DARK).text((r.incidencia?.descripcion||'—').slice(0,80), cols[1]+4, y+5, { width: colW[1]-8, ellipsis:true });
      doc.font('Helvetica').fontSize(7).fillColor(DARK).text(TIPMAP[r.incidencia?.tipologia]||r.incidencia?.tipologia||'—', cols[2]+4, y+5, { width: colW[2]-8 });
      doc.font('Helvetica').fontSize(7).fillColor(DARK).text((r.anotacion||'—').slice(0,40), cols[3]+4, y+5, { width: colW[3]-8, ellipsis:true });
      doc.font('Helvetica').fontSize(7.5).fillColor(DARK).text(String((r.documentacion||[]).length), cols[4]+4, y+5, { width: colW[4]-8, align:'center' });
      y += rh;
    });

    drawFooter('Portada Cierre CFO');

    // ══ PÁGINAS POR INCIDENCIA ══
    const TIPOLOGIA_LABELS = {
      material:   'Incidencia con necesidad de material',
      equipo:     'Incidencia sin necesidad de material. Resuelve directamente el equipo',
      documental: 'Incidencia Documental. Actualizar Documentación'
    };

    for (const r of resoluciones) {
      const inc = r.incidencia || {};

      // Página detalle incidencia original
      doc.addPage();
      drawPageHeader(`Incidencia ${r.incidenciaNumero}`);
      let iy = 106;

      doc.rect(M, iy, CW, 28).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(12).fillColor(WHITE)
         .text(`INCIDENCIA ${r.incidenciaNumero} de ${resoluciones.length} — DETALLE`, M + 10, iy + 8, { width: CW - 20 });
      iy += 34;

      const tipLabel = TIPOLOGIA_LABELS[inc.tipologia] || inc.tipologia || '—';
      doc.rect(M, iy, CW, 24).fill('#fff3cd').strokeColor('#f0c040').lineWidth(0.8).stroke();
      doc.rect(M, iy, 5, 24).fill('#e6a800');
      doc.font('Helvetica-Bold').fontSize(8.5).fillColor('#7a4000')
         .text('TIPOLOGÍA: ' + tipLabel.toUpperCase(), M + 12, iy + 7, { width: CW - 20 });
      iy += 30;

      if (inc.descripcion) {
        doc.rect(M, iy, CW, 14).fill(LGRAY); doc.rect(M, iy, 4, 14).fill(BLUE);
        doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('DESCRIPCIÓN DE LA INCIDENCIA', M+10, iy+4, { width: CW-20 });
        iy += 14;
        const dh = Math.min(Math.max(28, Math.ceil(inc.descripcion.length/85)*14+10), 80);
        doc.rect(M, iy, CW, dh).strokeColor(BORDER).lineWidth(0.4).stroke();
        doc.font('Helvetica').fontSize(9).fillColor(DARK).text(inc.descripcion, M+8, iy+7, { width: CW-16, height: dh-10, lineGap:2 });
        iy += dh + 6;
      }
      if (inc.descAdicional) {
        doc.rect(M, iy, CW, 14).fill(LGRAY); doc.rect(M, iy, 4, 14).fill('#2d8cd4');
        doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('INFORMACIÓN ADICIONAL', M+10, iy+4, { width: CW-20 });
        iy += 14;
        const dah = Math.min(Math.max(24, Math.ceil(inc.descAdicional.length/85)*14+10), 60);
        doc.rect(M, iy, CW, dah).strokeColor(BORDER).lineWidth(0.4).stroke();
        doc.font('Helvetica').fontSize(9).fillColor(DARK).text(inc.descAdicional, M+8, iy+7, { width: CW-16, height: dah-10, lineGap:2 });
        iy += dah + 6;
      }

      const fotosOriginales = [...(inc.fotosDoc||[]), ...(inc.infAdicional||[])].filter(f => f.dataUrl?.startsWith('data:image'));
      if (fotosOriginales.length > 0) {
        if (iy > H-120) { drawFooter(`Incidencia ${r.incidenciaNumero}`); doc.addPage(); drawPageHeader(`Incidencia ${r.incidenciaNumero}`); iy=106; }
        doc.rect(M, iy, CW, 16).fill(LGRAY); doc.rect(M, iy, 4, 16).fill(BLUE);
        doc.font('Helvetica-Bold').fontSize(7.5).fillColor(TEXTGRAY)
           .text(`FOTOGRAFÍAS ORIGINALES DE LA INCIDENCIA (${fotosOriginales.length})`, M+10, iy+5, { width: CW-20 });
        iy += 20;
        iy = insertarFotos(doc, fotosOriginales, M, iy, CW, W, H, DARK, BLUE, BORDER, WHITE,
          drawPageHeader, drawFooter, `Inc. ${r.incidenciaNumero} — Original`, pedido, fechaHora, operario);
      }

      // Página resolución
      if (iy > H-160) { drawFooter(`Incidencia ${r.incidenciaNumero}`); doc.addPage(); drawPageHeader(`Inc. ${r.incidenciaNumero} — Resolución`); iy=106; } else { iy+=12; }

      doc.rect(M, iy, CW, 28).fill(GREEN);
      doc.font('Helvetica-Bold').fontSize(12).fillColor(WHITE)
         .text(`RESOLUCIÓN — INCIDENCIA ${r.incidenciaNumero}`, M+10, iy+8, { width: CW-20 });
      iy += 34;

      fieldBox('Fecha de resolución', r.fechaResolucion || fechaHora, M, iy, CW, 30);
      iy += 36;

      if (r.anotacion) {
        doc.rect(M, iy, CW, 14).fill('#e2f4ea'); doc.rect(M, iy, 4, 14).fill(GREEN);
        doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('ANOTACIÓN ESPECIAL SOBRE LA RESOLUCIÓN', M+10, iy+4, { width: CW-20 });
        iy += 14;
        const ah = Math.min(Math.max(28, Math.ceil(r.anotacion.length/85)*14+10), 100);
        doc.rect(M, iy, CW, ah).strokeColor(BORDER).lineWidth(0.4).stroke();
        doc.font('Helvetica').fontSize(9).fillColor(DARK).text(r.anotacion, M+8, iy+7, { width: CW-16, height: ah-10, lineGap:2 });
        iy += ah + 8;
      }

      const fotosRes = (r.documentacion||[]).filter(f => f.dataUrl?.startsWith('data:image'));
      const docsRes  = (r.documentacion||[]).filter(f => f.dataUrl && !f.dataUrl.startsWith('data:image'));
      if (fotosRes.length > 0) {
        if (iy > H-120) { drawFooter(`Inc. ${r.incidenciaNumero} Res.`); doc.addPage(); drawPageHeader(`Inc. ${r.incidenciaNumero} — Resolución`); iy=106; }
        doc.rect(M, iy, CW, 16).fill('#e2f4ea'); doc.rect(M, iy, 4, 16).fill(GREEN);
        doc.font('Helvetica-Bold').fontSize(7.5).fillColor(TEXTGRAY)
           .text(`DOCUMENTACIÓN JUSTIFICATIVA DEL CIERRE (${fotosRes.length})`, M+10, iy+5, { width: CW-20 });
        iy += 20;
        iy = insertarFotos(doc, fotosRes, M, iy, CW, W, H, DARK, GREEN, BORDER, WHITE,
          drawPageHeader, drawFooter, `Inc. ${r.incidenciaNumero} — Resolución`, pedido, fechaHora, operario);
      }
      if (docsRes.length > 0) {
        if (iy > H-80) { drawFooter(`Inc. ${r.incidenciaNumero} Res.`); doc.addPage(); drawPageHeader(`Inc. ${r.incidenciaNumero} — Resolución`); iy=106; }
        doc.rect(M, iy, CW, 14).fill('#e2f4ea'); doc.rect(M, iy, 4, 14).fill(GREEN);
        doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY).text('DOCUMENTOS ADJUNTOS', M+10, iy+4, { width: CW-20 });
        iy += 18;
        docsRes.forEach(d => {
          if (iy > H-40) { drawFooter(`Inc. ${r.incidenciaNumero} Res.`); doc.addPage(); drawPageHeader(`Inc. ${r.incidenciaNumero} — Resolución`); iy=106; }
          doc.rect(M, iy, CW, 20).strokeColor(BORDER).lineWidth(0.3).stroke();
          doc.font('Helvetica').fontSize(8).fillColor(DARK).text('📄 '+(d.name||'documento'), M+8, iy+6, { width: CW-16 });
          iy += 22;
        });
      }

      drawFooter(`Incidencia ${r.incidenciaNumero} — Resolución`);
    }

    doc.end();
  });
}

app.post('/api/cfo', requireAuth, async (req, res) => {
  const { pedido, datosExcel, incidencias, timestamp } = req.body;
  if (!pedido)       return res.status(400).json({ success: false, error: 'Falta número de pedido' });
  if (!Array.isArray(incidencias) || incidencias.length === 0)
    return res.status(400).json({ success: false, error: 'No hay incidencias' });

  const operario  = req.user.name;
  const fechaHora = timestamp || new Date().toLocaleString('es-ES');
  const ts        = Date.now();

  // ── 1. Generar PDF ─────────────────────────────────────
  let pdfBuf;
  try {
    pdfBuf = await generarCFOPDF({ pedido, datosExcel, incidencias, fechaHora, operario });
  } catch (e) {
    return res.status(500).json({ success: false, error: 'Error generando PDF CFO: ' + e.message });
  }

  // ── 2. Generar JSON (CON imágenes en base64 para que Cierre Objeciones pueda mostrarlas) ──
  const jsonData = {
    pedido, datosExcel, fechaHora, operario,
    generadoEn: new Date().toISOString(),
    totalIncidencias: incidencias.length,
    incidencias: incidencias.map(inc => ({
      numero:        inc.numero,
      descripcion:   inc.descripcion   || '',
      tipologia:     inc.tipologia     || '',
      descAdicional: inc.descAdicional || '',
      fotosDoc:      inc.fotosDoc      || [],   // incluye dataUrl para visualización
      infAdicional:  inc.infAdicional  || [],   // incluye dataUrl para visualización
    }))
  };
  const jsonBuf = Buffer.from(JSON.stringify(jsonData, null, 2), 'utf8');

  // ── 3. Subir al FTP ────────────────────────────────────
  const client = new ftp.Client(300000);
  client.ftp.verbose = false;
  try {
    await client.access(FTP_CONFIG);

    // PDF → www/DMA/Apertura Objeciones  (PDF directo)
    const dirPDF  = path.posix.join(BASE_PATH, 'Apertura Objeciones');
    await ensureDir(client, dirPDF);
    const tsCFO   = tsNombre(req.user.id || '');
    const pdfName = `${String(pedido).trim()}_${tsCFO}_Informe_Visita_CFO.pdf`;
    await client.uploadFrom(Readable.from(pdfBuf), path.posix.join(dirPDF, pdfName));

    // JSON → www/DMA/Apertura Objeciones json
    const dirJSON  = path.posix.join(BASE_PATH, 'Apertura Objeciones json');
    await ensureDir(client, dirJSON);
    const jsonName = `${String(pedido).trim()} Informe Visita CFO_${tsCFO}.json`;
    await client.uploadFrom(Readable.from(jsonBuf), path.posix.join(dirJSON, jsonName));

    res.json({ success: true, pdfFile: pdfName, jsonFile: jsonName,
      rutaPDF:  `${dirPDF}/${pdfName}`,
      rutaJSON: `${dirJSON}/${jsonName}` });
  } catch (e) {
    res.status(500).json({ success: false, error: 'Error subiendo CFO: ' + e.message });
  } finally {
    client.close();
  }
});

// ── GENERADOR PDF CFO ──────────────────────────────────────────
async function generarCFOPDF({ pedido, datosExcel, incidencias, fechaHora, operario }) {
  return new Promise((resolve, reject) => {
    const doc  = new PDFDocument({ margin: 0, size: 'A4', autoFirstPage: true });
    const bufs = [];
    doc.on('data',  c => bufs.push(c));
    doc.on('end',   () => resolve(Buffer.concat(bufs)));
    doc.on('error', reject);

    const W  = 595, H = 842, M = 32;
    const CW = W - M * 2;
    const BLUE     = '#1B6CA8';
    const DARK     = '#1a2332';
    const BORDER   = '#d0d8e4';
    const TEXTGRAY = '#5a6478';
    const WHITE    = '#ffffff';
    const LGRAY    = '#f4f7fb';

    // Logo buffer
    let logoBuf = null;
    if (LOGO_B64) {
      try { logoBuf = Buffer.from(LOGO_B64.replace(/^data:image\/\w+;base64,/, ''), 'base64'); } catch (_) {}
    }

    const dx = datosExcel || {};

    // ── helpers ──────────────────────────────────────────
    const drawPageHeader = (label) => {
      doc.rect(0, 0, W, 90).fill(DARK);
      if (logoBuf) { try { doc.image(logoBuf, M, 10, { height: 68, fit: [68, 68] }); } catch (_) {} }
      doc.font('Helvetica-Bold').fontSize(15).fillColor(WHITE)
         .text('ROURA & CEVASA', M + 78, 14, { width: 220 });
      doc.font('Helvetica').fontSize(8).fillColor('#a8c4e0')
         .text('Visita CFO — Apertura de Objeciones', M + 78, 34, { width: 220 });
      doc.font('Helvetica').fontSize(7.5).fillColor('#7a9abf')
         .text(fechaHora, M + 78, 50, { width: 220 });
      doc.roundedRect(W - M - 132, 10, 132, 68, 6).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(9).fillColor(WHITE)
         .text(label.toUpperCase(), W - M - 132, 28, { width: 132, align: 'center' });
      doc.font('Helvetica').fontSize(7).fillColor('#cde4f5')
         .text('Pedido: ' + pedido, W - M - 132, 50, { width: 132, align: 'center' });
      doc.rect(0, 90, W, 4).fill(BLUE);
    };

    const drawFooter = (pageLabel) => {
      const py = H - 26;
      doc.rect(0, py, W, 26).fill(DARK);
      doc.font('Helvetica').fontSize(7).fillColor('#7a9abf')
         .text(`${pageLabel}  ·  Pedido: ${pedido}  ·  Operario: ${operario}  ·  ${fechaHora}`,
               M, py + 8, { width: CW, align: 'center' });
    };

    const fieldBox = (label, value, x, y, w, h = 34) => {
      doc.rect(x, y, w, h).strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.rect(x, y, w, 12).fill(LGRAY);
      doc.font('Helvetica-Bold').fontSize(6.5).fillColor(TEXTGRAY)
         .text(label.toUpperCase(), x + 5, y + 3, { width: w - 10 });
      doc.font('Helvetica').fontSize(9).fillColor(DARK)
         .text(String(value || '—'), x + 5, y + 15, { width: w - 10, height: h - 17, ellipsis: true });
    };

    // ══════════════════════════════════════════════════════
    // PÁGINA 1 — PORTADA / DATOS INSTALACIÓN
    // ══════════════════════════════════════════════════════
    drawPageHeader('CFO');

    let y = 106;

    // Banner título
    doc.rect(M, y, CW, 30).fill(BLUE);
    doc.font('Helvetica-Bold').fontSize(13).fillColor(WHITE)
       .text('VISITA CFO — APERTURA DE OBJECIONES', M + 10, y + 9, { width: CW - 20 });
    y += 38;

    // Sub-header instalación
    doc.rect(M, y, CW, 26).fill(LGRAY);
    doc.rect(M, y, 4, 26).fill(BLUE);
    doc.font('Helvetica-Bold').fontSize(10).fillColor(DARK)
       .text('Datos de la Instalación', M + 12, y + 8, { width: CW - 20 });
    y += 32;

    // Grid de datos
    const half = CW / 2;
    fieldBox('Número de Pedido', pedido,          M,        y, half,      34);
    fieldBox('SIE',              dx.SIE || '—',   M + half, y, half,      34);
    y += 38;
    fieldBox('Referencia Cliente', dx.RefCli || '—', M,        y, half,   34);
    fieldBox('PA',                 dx.PA    || '—',  M + half, y, half,   34);
    y += 38;
    fieldBox('Ciudad',      dx.Ciudad    || '—', M,           y, CW*0.4,  34);
    fieldBox('Provincia',   dx.Provincia || '—', M + CW*0.4,  y, CW*0.35, 34);
    fieldBox('Fecha',       fechaHora,           M + CW*0.75, y, CW*0.25, 34);
    y += 38;
    fieldBox('Dirección',   dx.Direccion  || '—', M, y, CW, 34);
    y += 38;
    fieldBox('Descripción', dx.Descripcion|| '—', M, y, CW, 34);
    y += 38;
    fieldBox('Operario',    operario || '—',      M, y, CW, 30);
    y += 36;

    // Separador
    doc.rect(M, y + 6, CW, 1).fill(BORDER); y += 16;

    // Resumen incidencias
    doc.rect(M, y, CW, 26).fill(BLUE);
    doc.font('Helvetica-Bold').fontSize(10).fillColor(WHITE)
       .text(`Resumen: ${incidencias.length} incidencia(s) registrada(s)`, M + 10, y + 8, { width: CW - 20 });
    y += 32;

    // Tabla resumen
    const colW = [40, 230, 185, 60];
    const cols = [M, M+colW[0], M+colW[0]+colW[1], M+colW[0]+colW[1]+colW[2]];
    // Header tabla
    ['Nº', 'Descripción', 'Tipología', 'Fotos'].forEach((h, i) => {
      doc.rect(cols[i], y, colW[i], 20).fill('#e2eaf4').strokeColor(BORDER).lineWidth(0.5).stroke();
      doc.font('Helvetica-Bold').fontSize(7.5).fillColor(DARK)
         .text(h, cols[i] + 4, y + 6, { width: colW[i] - 8 });
    });
    y += 20;

    const TIPMAP = {
      'material':   'Nec. Material',
      'equipo':     'Equipo resuelve',
      'documental': 'Documental'
    };
    incidencias.slice(0, 40).forEach(inc => {
      const rh = 18;
      if (y + rh > H - 40) { drawFooter('Portada CFO'); doc.addPage(); drawPageHeader('CFO — Resumen'); y = 106; }
      const rowBg = (inc.numero % 2 === 0) ? '#f8fafc' : WHITE;
      [0,1,2,3].forEach(i => doc.rect(cols[i], y, colW[i], rh).fill(rowBg).strokeColor(BORDER).lineWidth(0.3).stroke());
      doc.font('Helvetica').fontSize(7.5).fillColor(DARK)
         .text(String(inc.numero), cols[0]+4, y+5, { width: colW[0]-8 });
      doc.font('Helvetica').fontSize(7).fillColor(DARK)
         .text((inc.descripcion||'—').slice(0,80), cols[1]+4, y+5, { width: colW[1]-8, ellipsis:true });
      doc.font('Helvetica').fontSize(7).fillColor(DARK)
         .text(TIPMAP[inc.tipologia]||inc.tipologia||'—', cols[2]+4, y+5, { width: colW[2]-8 });
      const nFotos = (inc.fotosDoc||[]).length + (inc.infAdicional||[]).length;
      doc.font('Helvetica').fontSize(7.5).fillColor(DARK)
         .text(String(nFotos), cols[3]+4, y+5, { width: colW[3]-8, align:'center' });
      y += rh;
    });

    drawFooter('Portada CFO');

    // ══════════════════════════════════════════════════════
    // PÁGINAS DE INCIDENCIAS
    // ══════════════════════════════════════════════════════
    const TIPOLOGIA_LABELS = {
      'material':   'Incidencia con necesidad de material',
      'equipo':     'Incidencia sin necesidad de material. Resuelve directamente el equipo',
      'documental': 'Incidencia Documental. Actualizar Documentación'
    };

    for (const inc of incidencias) {
      doc.addPage();
      drawPageHeader(`Incidencia ${inc.numero}`);
      let iy = 106;

      // Banner incidencia
      doc.rect(M, iy, CW, 28).fill(BLUE);
      doc.font('Helvetica-Bold').fontSize(12).fillColor(WHITE)
         .text(`INCIDENCIA ${inc.numero} de ${incidencias.length}`, M + 10, iy + 8, { width: CW - 20 });
      iy += 34;

      // Tipología
      const tipLabel = TIPOLOGIA_LABELS[inc.tipologia] || inc.tipologia || '—';
      doc.rect(M, iy, CW, 24).fill('#fff3cd').strokeColor('#f0c040').lineWidth(0.8).stroke();
      doc.rect(M, iy, 5, 24).fill('#e6a800');
      doc.font('Helvetica-Bold').fontSize(8.5).fillColor('#7a4000')
         .text('TIPOLOGÍA: ' + tipLabel.toUpperCase(), M + 12, iy + 7, { width: CW - 20 });
      iy += 30;

      // Descripción incidencia
      if (inc.descripcion) {
        doc.rect(M, iy, CW, 14).fill(LGRAY);
        doc.rect(M, iy, 4, 14).fill(BLUE);
        doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY)
           .text('DESCRIPCIÓN DE LA INCIDENCIA', M + 10, iy + 4, { width: CW - 20 });
        iy += 14;
        const descH = Math.min(Math.max(28, Math.ceil(inc.descripcion.length / 85) * 14 + 10), 80);
        doc.rect(M, iy, CW, descH).strokeColor(BORDER).lineWidth(0.4).stroke();
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
           .text(inc.descripcion, M + 8, iy + 7, { width: CW - 16, height: descH - 10, lineGap: 2 });
        iy += descH + 6;
      }

      // Descripción adicional
      if (inc.descAdicional) {
        doc.rect(M, iy, CW, 14).fill(LGRAY);
        doc.rect(M, iy, 4, 14).fill('#2d8cd4');
        doc.font('Helvetica-Bold').fontSize(7).fillColor(TEXTGRAY)
           .text('DESCRIPCIÓN ADICIONAL', M + 10, iy + 4, { width: CW - 20 });
        iy += 14;
        const daH = Math.min(Math.max(24, Math.ceil(inc.descAdicional.length / 85) * 14 + 10), 70);
        doc.rect(M, iy, CW, daH).strokeColor(BORDER).lineWidth(0.4).stroke();
        doc.font('Helvetica').fontSize(9).fillColor(DARK)
           .text(inc.descAdicional, M + 8, iy + 7, { width: CW - 16, height: daH - 10, lineGap: 2 });
        iy += daH + 6;
      }

      // Fotos documentación principal (campo 1)
      const fotosPrinc = (inc.fotosDoc || []).filter(f => f.dataUrl?.startsWith('data:image'));
      if (fotosPrinc.length > 0) {
        doc.rect(M, iy, CW, 16).fill(LGRAY);
        doc.rect(M, iy, 4, 16).fill(BLUE);
        doc.font('Helvetica-Bold').fontSize(7.5).fillColor(TEXTGRAY)
           .text(`DOCUMENTACIÓN / FOTOGRAFÍAS (${fotosPrinc.length})`, M + 10, iy + 5, { width: CW - 20 });
        iy += 20;
        iy = insertarFotos(doc, fotosPrinc, M, iy, CW, W, H, DARK, BLUE, BORDER, WHITE, drawPageHeader, drawFooter, `Inc. ${inc.numero}`, pedido, fechaHora, operario);
      }

      // Fotos información adicional (campo 4)
      const fotosAd = (inc.infAdicional || []).filter(f => f.dataUrl?.startsWith('data:image'));
      if (fotosAd.length > 0) {
        if (iy > H - 120) { drawFooter(`Incidencia ${inc.numero}`); doc.addPage(); drawPageHeader(`Incidencia ${inc.numero}`); iy = 106; }
        doc.rect(M, iy, CW, 16).fill('#e8f2fb');
        doc.rect(M, iy, 4, 16).fill('#2d8cd4');
        doc.font('Helvetica-Bold').fontSize(7.5).fillColor(TEXTGRAY)
           .text(`INFORMACIÓN ADICIONAL / CROQUIS (${fotosAd.length})`, M + 10, iy + 5, { width: CW - 20 });
        iy += 20;
        iy = insertarFotos(doc, fotosAd, M, iy, CW, W, H, DARK, BLUE, BORDER, WHITE, drawPageHeader, drawFooter, `Inc. ${inc.numero} — Adicional`, pedido, fechaHora, operario);
      }

      drawFooter(`Incidencia ${inc.numero} de ${incidencias.length}`);
    }

    doc.end();
  });
}

// Inserta fotos en grid 2 columnas, añade páginas si es necesario
function insertarFotos(doc, fotos, M, startY, CW, W, H, DARK, BLUE, BORDER, WHITE, drawPageHeader, drawFooter, secLabel, pedido, fechaHora, operario) {
  const GAP    = 8;
  const fotoW  = (CW - GAP) / 2;
  const fotoH  = fotoW * 0.72;  // aspecto 4:3 aprox
  let iy = startY;
  let col = 0;

  for (let i = 0; i < fotos.length; i++) {
    // Nueva fila: si no cabe, nueva página
    if (col === 0 && iy + fotoH > H - 50) {
      drawFooter(secLabel);
      doc.addPage();
      drawPageHeader(secLabel);
      iy = 106;
    }

    const fx = col === 0 ? M : M + fotoW + GAP;
    const fy = iy;

    // Marco foto
    doc.rect(fx, fy, fotoW, fotoH).fill('#f0f4f8').strokeColor(BORDER).lineWidth(0.5).stroke();

    try {
      const b64  = fotos[i].dataUrl.replace(/^data:image\/\w+;base64,/, '');
      const imgB = Buffer.from(b64, 'base64');
      doc.image(imgB, fx + 3, fy + 3, {
        width:  fotoW - 6,
        height: fotoH - 6,
        fit:    [fotoW - 6, fotoH - 6],
        align:  'center',
        valign: 'center'
      });
    } catch (_) {
      doc.font('Helvetica').fontSize(8).fillColor(DARK)
         .text('Imagen no disponible', fx + 6, fy + fotoH / 2 - 6, { width: fotoW - 12, align: 'center' });
    }

    // Nombre archivo
    doc.font('Helvetica').fontSize(6.5).fillColor('#6a7a8a')
       .text(fotos[i].name || `foto_${i+1}`, fx + 3, fy + fotoH - 10, { width: fotoW - 6, ellipsis: true });

    col++;
    if (col === 2) { col = 0; iy += fotoH + GAP; }
  }

  // Si quedó columna suelta
  if (col === 1) iy += fotoH + GAP;
  return iy;
}



// ── EXPEDICIONES ─────────────────────────────────────────────────────────────
app.post('/api/expediciones/fotos', requireAuth, async (req, res) => {
  const { pedido, fotos } = req.body;
  if (!pedido || !Array.isArray(fotos) || fotos.length === 0)
    return res.status(400).json({ success: false, error: 'Faltan datos (pedido o fotos)' });

  const tDir = path.posix.join(BASE_PATH, 'Expediciones');
  const entries = [];
  for (const f of fotos) {
    try {
      if (!f.dataUrl || !f.dataUrl.startsWith('data:image')) continue;
      const b64    = f.dataUrl.replace(/^data:image\/[^;]+;base64,/, '');
      const buffer = Buffer.from(b64, 'base64');
      const ext    = (f.type || 'image/jpeg').split('/')[1] || 'jpg';
      entries.push({ name: f.name || ('foto_' + Date.now() + '.' + ext), buffer });
    } catch (_) {}
  }
  if (entries.length === 0)
    return res.status(400).json({ success: false, error: 'Sin imágenes válidas' });

  const client = new ftp.Client(180000);
  client.ftp.verbose = false;
  try {
    const zipBuffer = await crearZipBuffer(entries);
    const ts        = tsNombre(req.user.id || '');
    const zipName   = `${String(pedido).trim()}_${ts}_Fotografias_Expediciones.zip`;
    await client.access(FTP_CONFIG);
    await ensureDir(client, tDir);
    await client.uploadFrom(Readable.from(zipBuffer), path.posix.join(tDir, zipName));
    res.json({ success: true, zipFile: zipName, count: entries.length });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  } finally {
    client.close();
  }
});

// ── SUBIDA DE FOTOS BASE64 (CFO / Cierre / Mod) ───────────────────────────
app.post('/api/upload-fotos-b64', requireAuth, async (req, res) => {
  const { pedido, categoria, fotos } = req.body;
  if (!pedido || !Array.isArray(fotos) || fotos.length === 0)
    return res.status(400).json({ success: false, error: 'Faltan datos' });

  const catMap = {
    'fotos_medicion':   { dir: 'TD',           label: 'Fotografias_Medicion' },
    'fotos_antes':      { dir: 'Fotos Inicio', label: 'Fotografias_Antes' },
    'fotos_final':      { dir: 'Fotos Fin',    label: 'Fotografias_Final' },
    'fotos_cfo':        { dir: 'Fotos Fin',    label: 'Fotografias_Visita CFO' },
    'fotos_cierre_cfo': { dir: 'Fotos Fin',    label: 'Fotografias_Cierre CFO' },
  };
  const catStr  = String(categoria).trim();
  const catInfo = catMap[catStr] || { dir: categoria, label: categoria };
  const tDir    = path.posix.join(BASE_PATH, catInfo.dir);

  // Las fotos antes/final también se copian a la carpeta de espera de certificación
  const esFotosCertificacion = (catStr === 'fotos_antes' || catStr === 'fotos_final');
  const dirPteCert = path.posix.join(BASE_PATH, 'Fotografias Pte Certificacion');

  const entries = [];
  for (const f of fotos) {
    try {
      if (!f.dataUrl || !f.dataUrl.startsWith('data:image')) continue;
      const b64    = f.dataUrl.replace(/^data:image\/[^;]+;base64,/, '');
      const buffer = Buffer.from(b64, 'base64');
      const ext    = (f.type || 'image/jpeg').split('/')[1] || 'jpg';
      entries.push({ name: f.name || ('foto_' + Date.now() + '.' + ext), buffer });
    } catch (_) {}
  }
  if (entries.length === 0)
    return res.status(400).json({ success: false, error: 'Sin imágenes válidas' });

  const client = new ftp.Client(180000);
  client.ftp.verbose = false;
  try {
    const zipBuffer = await crearZipBuffer(entries);
    const ts        = tsNombre(req.user.id || '');
    const zipName   = `${String(pedido).trim()}_${ts}_${catInfo.label}.zip`;
    await client.access(FTP_CONFIG);
    await ensureDir(client, tDir);
    await client.uploadFrom(Readable.from(zipBuffer), path.posix.join(tDir, zipName));

    // Si son fotos de antes o final, subir también a "Fotografias Pte Certificacion"
    if (esFotosCertificacion) {
      try {
        await ensureDir(client, dirPteCert);
        await client.uploadFrom(Readable.from(zipBuffer), path.posix.join(dirPteCert, zipName));
      } catch (eCert) {
        console.warn('[upload-fotos-b64] No se pudo copiar a Fotografias Pte Certificacion:', eCert.message);
      }
    }

    res.json({ success: true, zipFile: zipName, count: entries.length });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  } finally {
    client.close();
  }
});

// ── INSPECCIONES DE SEGURIDAD — Fotos ────────────────────────────────────────
app.post('/api/inspecciones-seguridad/fotos', requireAuth, upload.array('files', 20), async (req, res) => {
  const { pedido } = req.body;
  if (!pedido) return res.status(400).json({ success: false, error: 'Falta número de pedido' });
  if (!req.files || req.files.length === 0) return res.status(400).json({ success: false, error: 'No se recibieron archivos' });

  const tDir = path.posix.join(BASE_PATH, 'Inspecciones Seguridad');
  const entries = req.files.map(file => {
    const ext  = path.extname(file.originalname) || '';
    const base = path.basename(file.originalname, ext).replace(/[^a-zA-Z0-9_-]/g, '_').slice(0, 60);
    const safe = Date.now() + '_' + base + ext;
    return { name: safe, buffer: file.buffer };
  });

  const client = new ftp.Client(180000);
  client.ftp.verbose = false;
  try {
    const zipBuffer = await crearZipBuffer(entries);
    const ts        = tsNombre(req.user.id || '');
    const zipName   = `${String(pedido).trim()}_${ts}_Fotografias_Inspeccion_Seguridad.zip`;
    await client.access(FTP_CONFIG);
    await ensureDir(client, tDir);
    await client.uploadFrom(Readable.from(zipBuffer), path.posix.join(tDir, zipName));
    res.json({ success: true, zipFile: zipName, count: entries.length });
  } catch (e) {
    res.status(500).json({ success: false, error: e.message });
  } finally {
    client.close();
  }
});

// ── GLOBAL ERROR HANDLER (siempre devuelve JSON) ─────────────
app.use((err, req, res, next) => {
  console.error('Error no controlado:', err.message);
  const status = err.status || err.statusCode || 500;
  res.status(status).json({ success: false, error: err.message || 'Error interno del servidor' });
});

// ── START ────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  getUsers();
  console.log(`✅ FTP Manager v4.2 en puerto ${PORT}`);
  console.log(`👤 Admin: admin / admin123`);
  console.log(`📁 FTP base: ${BASE_PATH} (FTPS)`);
  console.log(`🖼  Logo: ${LOGO_B64 ? 'OK' : 'No encontrado (usando fallback)'}`);
});
