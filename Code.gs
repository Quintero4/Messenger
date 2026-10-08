/**
 * RetroTalk 2000 - Chat tipo Messenger sobre Google Sheets
 *
 * DESPLIEGUE: Implementar > Aplicación web
 *   - Ejecutar como: "Usuario que accede a la aplicación web"
 *   - Quién tiene acceso: "Cualquier usuario de tu dominio" (Workspace)
 * El HTML debe llamarse exactamente "Index" (Index.html).
 */
const APP = {
  USERS_SHEET: 'Usuarios',
  MESSAGES_SHEET: 'Mensajes',
  ONLINE_MS: 45000,   // sin actividad durante este tiempo => desconectado
  TOUCH_MS: 15000,    // como máximo se escribe la actividad cada 15 s por usuario
  MAX_ROWS: 500,      // filas recientes de Mensajes que se leen en cada consulta
  MAX_LEN: 1000,
  NUDGE_MS: 10000,    // tiempo mínimo entre zumbidos del mismo usuario
  STATES: ['disponible', 'llamada', 'cafe', 'desconectado'],
  SPREADSHEET_ID: ''  // Vacío si el script está vinculado a la hoja; si no, pega aquí el ID.
};

/* ---------- Utilidades ---------- */

function ss_() {
  const ss = APP.SPREADSHEET_ID
    ? SpreadsheetApp.openById(APP.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No se encontró la hoja de cálculo. Vincula el script a la hoja o rellena APP.SPREADSHEET_ID.');
  return ss;
}

function sheet_(name) {
  const sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('Falta la hoja "' + name + '".');
  return sheet;
}

const norm_ = v => String(v || '').trim().toLowerCase();

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('RetroTalk 2000')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setupDatabase_() {
  const ss = ss_();
  if (ss.getSheetByName(APP.USERS_SHEET) && ss.getSheetByName(APP.MESSAGES_SHEET)) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!ss.getSheetByName(APP.USERS_SHEET)) {
      const users = ss.insertSheet(APP.USERS_SHEET);
      users.getRange('A:B').setNumberFormat('@');
      users.getRange('C:C').setNumberFormat('yyyy-mm-dd hh:mm:ss');
      users.getRange('D:D').setNumberFormat('@');
      users.appendRow(['email', 'nombre', 'ultimaActividad', 'estado']);
      users.setFrozenRows(1);
    }
    if (!ss.getSheetByName(APP.MESSAGES_SHEET)) {
      const messages = ss.insertSheet(APP.MESSAGES_SHEET);
      messages.getRange('A:A').setNumberFormat('@');
      messages.getRange('C:G').setNumberFormat('@');
      messages.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss.000');
      messages.appendRow(['id', 'fecha', 'emisorEmail', 'emisorNombre', 'receptorEmail', 'mensaje', 'tipo']);
      messages.setFrozenRows(1);
    }
  } finally {
    lock.releaseLock();
  }
}

/** Migra hojas creadas con la versión anterior (añade columnas "estado" y "tipo"). */
function ensureSchema_() {
  const cache = CacheService.getScriptCache();
  if (cache.get('schema_ok')) return;
  const users = sheet_(APP.USERS_SHEET);
  const messages = sheet_(APP.MESSAGES_SHEET);
  if (!users.getRange('D1').getValue()) {
    users.getRange('D:D').setNumberFormat('@');
    users.getRange('D1').setValue('estado');
  }
  if (!messages.getRange('G1').getValue()) {
    messages.getRange('G:G').setNumberFormat('@');
    messages.getRange('G1').setValue('tipo');
  }
  cache.put('schema_ok', '1', 21600);
}

function getIdentity_() {
  const email = norm_(Session.getActiveUser().getEmail());
  if (!email) {
    throw new Error('No se pudo identificar al usuario. Publica la Web App para usuarios de tu dominio y ejecútala como "usuario que accede".');
  }
  const nombre = email.split('@')[0]
    .replace(/[._-]+/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase());
  return { email, nombre };
}

function findUserRow_(sheet, email) {
  const last = sheet.getLastRow();
  if (last < 2) return -1;
  const emails = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < emails.length; i++) {
    if (norm_(emails[i][0]) === email) return i + 2;
  }
  return -1;
}

/** Registra actividad (limitado por caché para no escribir en la hoja cada pocos segundos). */
function touchUser_(user) {
  const cache = CacheService.getScriptCache();
  const key = 'touch_' + user.email;
  if (cache.get(key)) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (cache.get(key)) return;
    const sheet = sheet_(APP.USERS_SHEET);
    const row = findUserRow_(sheet, user.email);
    if (row === -1) {
      sheet.appendRow([user.email, user.nombre, new Date(), 'disponible']);
    } else {
      sheet.getRange(row, 2, 1, 2).setValues([[user.nombre, new Date()]]); // no toca la columna "estado"
    }
    cache.put(key, '1', Math.max(1, Math.floor(APP.TOUCH_MS / 1000)));
  } finally {
    lock.releaseLock();
  }
}

function getMyState_(email) {
  const sheet = sheet_(APP.USERS_SHEET);
  const row = findUserRow_(sheet, email);
  if (row === -1) return 'disponible';
  const value = String(sheet.getRange(row, 4).getValue() || '');
  return APP.STATES.indexOf(value) >= 0 ? value : 'disponible';
}

function getContacts_(me) {
  const sheet = sheet_(APP.USERS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const cutoff = Date.now() - APP.ONLINE_MS;
  return sheet.getRange(2, 1, last - 1, 4).getValues()
    .filter(row => row[0] && norm_(row[0]) !== me.email)
    .map(row => {
      const active = new Date(row[2]).getTime() >= cutoff;
      const saved = APP.STATES.indexOf(row[3]) >= 0 ? row[3] : 'disponible';
      const estado = active ? saved : 'desconectado';   // sin actividad => desconectado
      return {
        email: norm_(row[0]),
        nombre: String(row[1] || row[0]),
        estado: estado,
        online: estado !== 'desconectado'
      };
    })
    .sort((a, b) => Number(b.online) - Number(a.online) || a.nombre.localeCompare(b.nombre));
}

/** Lee las últimas MAX_ROWS filas de mensajes (o solo desde minRow si es mayor). */
function readMessages_(minRow) {
  const sheet = sheet_(APP.MESSAGES_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { values: [], start: 2, lastRow: lastRow };
  const start = Math.max(2, lastRow - APP.MAX_ROWS + 1, minRow || 2);
  if (start > lastRow) return { values: [], start: start, lastRow: lastRow };
  return {
    values: sheet.getRange(start, 1, lastRow - start + 1, 7).getValues(),
    start: start,
    lastRow: lastRow
  };
}

function toMessage_(row, rowNumber) {
  return {
    fila: rowNumber,
    id: String(row[0]),
    fecha: new Date(row[1]).toISOString(),
    emisorEmail: norm_(row[2]),
    emisorNombre: String(row[3]),
    receptorEmail: norm_(row[4]),
    mensaje: String(row[5]),
    tipo: row[6] === 'zumbido' ? 'zumbido' : 'texto'
  };
}

/* ---------- API pública (google.script.run) ---------- */

function bootstrap() {
  setupDatabase_();
  ensureSchema_();
  const me = getIdentity_();
  touchUser_(me);
  me.estado = getMyState_(me.email);
  return { me: me, contactos: getContacts_(me) };
}

/**
 * Una sola llamada por ciclo: actividad + contactos + mensajes nuevos + avisos entrantes.
 * @param {string} contactEmail contacto abierto ('' si ninguno)
 * @param {number} afterRow     última fila ya recibida para ese chat (0 = carga inicial)
 */
function sync(contactEmail, afterRow) {
  const me = getIdentity_();
  touchUser_(me);
  const contactos = getContacts_(me);
  const contact = norm_(contactEmail);
  const after = Number(afterRow) || 0;

  const data = readMessages_(after > 0 ? after + 1 : 2);
  const incoming = {};   // email -> última fila que me ha enviado (cualquier tipo)
  const nudges = {};     // email -> última fila de zumbido que me ha enviado
  const mensajes = [];

  const scan = after > 0 ? readMessages_(Math.max(2, data.lastRow - 100)) : data;
  scan.values.forEach((row, i) => {
    if (norm_(row[4]) === me.email) {
      const sender = norm_(row[2]), n = scan.start + i;
      incoming[sender] = n;
      if (row[6] === 'zumbido') nudges[sender] = n;
    }
  });

  data.values.forEach((row, i) => {
    const n = data.start + i;
    if (n <= after || !contact) return;
    const s = norm_(row[2]), r = norm_(row[4]);
    if ((s === me.email && r === contact) || (s === contact && r === me.email)) {
      mensajes.push(toMessage_(row, n));
    }
  });

  return { contactos: contactos, mensajes: mensajes, incoming: incoming, nudges: nudges, ultimaFila: data.lastRow };
}

function getConversation(contactEmail, afterRow) {
  return sync(contactEmail, afterRow).mensajes;
}

function heartbeat() {
  const me = getIdentity_();
  touchUser_(me);
  return getContacts_(me);
}

/** Cambia el estado: disponible | llamada | cafe | desconectado (aparecer desconectado). */
function setStatus(estado) {
  const me = getIdentity_();
  if (APP.STATES.indexOf(estado) < 0) throw new Error('Estado no válido.');
  touchUser_(me);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = sheet_(APP.USERS_SHEET);
    const row = findUserRow_(sheet, me.email);
    if (row > 0) sheet.getRange(row, 4).setValue(estado);
  } finally {
    lock.releaseLock();
  }
  return estado;
}

/** El cliente la llama al salir de la página o cuando el equipo se bloquea / la pestaña queda oculta. */
function goOffline() {
  const me = getIdentity_();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = sheet_(APP.USERS_SHEET);
    const row = findUserRow_(sheet, me.email);
    if (row > 0) sheet.getRange(row, 3).setValue(new Date(0)); // actividad muy antigua => desconectado
    CacheService.getScriptCache().remove('touch_' + me.email);
  } finally {
    lock.releaseLock();
  }
}

function append_(me, to, text, tipo) {
  if (!to) throw new Error('Selecciona un contacto.');
  if (to === me.email) throw new Error('No puedes enviarte mensajes a ti mismo.');
  if (!getContacts_(me).some(c => c.email === to)) throw new Error('El contacto no existe.');
  touchUser_(me);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const sheet = sheet_(APP.MESSAGES_SHEET);
    const rowNumber = sheet.getLastRow() + 1;
    const item = {
      fila: rowNumber,
      id: Utilities.getUuid(),
      fecha: now.toISOString(),
      emisorEmail: me.email,
      emisorNombre: me.nombre,
      receptorEmail: to,
      mensaje: text,
      tipo: tipo
    };
    const range = sheet.getRange(rowNumber, 1, 1, 7);
    range.setNumberFormats([['@', 'yyyy-mm-dd hh:mm:ss.000', '@', '@', '@', '@', '@']]);
    range.setValues([[item.id, now, item.emisorEmail, item.emisorNombre, item.receptorEmail, item.mensaje, item.tipo]]);
    SpreadsheetApp.flush();
    return item;
  } finally {
    lock.releaseLock();
  }
}

function sendMessage(contactEmail, message) {
  const me = getIdentity_();
  const text = String(message || '').trim();
  if (!text) throw new Error('El mensaje está vacío.');
  if (text.length > APP.MAX_LEN) throw new Error('El mensaje supera los ' + APP.MAX_LEN + ' caracteres.');
  return append_(me, norm_(contactEmail), text, 'texto');
}

function sendNudge(contactEmail) {
  const me = getIdentity_();
  const cache = CacheService.getScriptCache();
  const key = 'nudge_' + me.email;
  if (cache.get(key)) throw new Error('Espera unos segundos antes de enviar otro zumbido.');
  const item = append_(me, norm_(contactEmail), '¡Zumbido!', 'zumbido');
  cache.put(key, '1', Math.max(1, Math.ceil(APP.NUDGE_MS / 1000)));
  return item;
}
