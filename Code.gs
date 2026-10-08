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
  ONLINE_MS: 45000,   // tiempo sin actividad para considerar "desconectado"
  TOUCH_MS: 15000,    // como máximo se escribe la actividad cada 15 s por usuario
  MAX_ROWS: 500,      // filas recientes de Mensajes que se leen en cada consulta
  MAX_LEN: 1000,
  SPREADSHEET_ID: ''  // Déjalo vacío si el script está vinculado a la hoja.
                      // Si el script es independiente, pega aquí el ID de la hoja.
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

  // Bloqueo para que dos usuarios simultáneos no creen las hojas dos veces
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!ss.getSheetByName(APP.USERS_SHEET)) {
      const users = ss.insertSheet(APP.USERS_SHEET);
      users.getRange('A:B').setNumberFormat('@');                      // texto plano
      users.getRange('C:C').setNumberFormat('yyyy-mm-dd hh:mm:ss');
      users.appendRow(['email', 'nombre', 'ultimaActividad']);
      users.setFrozenRows(1);
    }
    if (!ss.getSheetByName(APP.MESSAGES_SHEET)) {
      const messages = ss.insertSheet(APP.MESSAGES_SHEET);
      messages.getRange('A:A').setNumberFormat('@');
      messages.getRange('C:F').setNumberFormat('@');                   // evita fórmulas/fechas/números automáticos
      messages.getRange('B:B').setNumberFormat('yyyy-mm-dd hh:mm:ss.000');
      messages.appendRow(['id', 'fecha', 'emisorEmail', 'emisorNombre', 'receptorEmail', 'mensaje']);
      messages.setFrozenRows(1);
    }
  } finally {
    lock.releaseLock();
  }
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

/** Registra actividad. Limitado por caché para no escribir en la hoja cada pocos segundos. */
function touchUser_(user) {
  const cache = CacheService.getScriptCache();
  const key = 'touch_' + user.email;
  if (cache.get(key)) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (cache.get(key)) return;
    const sheet = sheet_(APP.USERS_SHEET);
    const last = sheet.getLastRow();
    let row = -1;
    if (last >= 2) {
      const emails = sheet.getRange(2, 1, last - 1, 1).getValues();
      for (let i = 0; i < emails.length; i++) {
        if (norm_(emails[i][0]) === user.email) { row = i + 2; break; }
      }
    }
    if (row === -1) {
      sheet.appendRow([user.email, user.nombre, new Date()]);
    } else {
      sheet.getRange(row, 2, 1, 2).setValues([[user.nombre, new Date()]]);
    }
    cache.put(key, '1', Math.max(1, Math.floor(APP.TOUCH_MS / 1000)));
  } finally {
    lock.releaseLock();
  }
}

function getContacts_(me) {
  const sheet = sheet_(APP.USERS_SHEET);
  const last = sheet.getLastRow();
  if (last < 2) return [];
  const cutoff = Date.now() - APP.ONLINE_MS;
  return sheet.getRange(2, 1, last - 1, 3).getValues()
    .filter(row => row[0] && norm_(row[0]) !== me.email)
    .map(row => ({
      email: norm_(row[0]),
      nombre: String(row[1] || row[0]),
      online: new Date(row[2]).getTime() >= cutoff
    }))
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
    values: sheet.getRange(start, 1, lastRow - start + 1, 6).getValues(),
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
    mensaje: String(row[5])
  };
}

/* ---------- API pública (google.script.run) ---------- */

function bootstrap() {
  setupDatabase_();
  const me = getIdentity_();
  touchUser_(me);
  return { me: me, contactos: getContacts_(me) };
}

/**
 * Una sola llamada por ciclo: actividad + contactos + mensajes nuevos + avisos de mensajes entrantes.
 * @param {string} contactEmail contacto abierto ('' si ninguno)
 * @param {number} afterRow     última fila de la hoja que el cliente ya recibió (0 = carga inicial)
 */
function sync(contactEmail, afterRow) {
  const me = getIdentity_();
  touchUser_(me);
  const contactos = getContacts_(me);
  const contact = norm_(contactEmail);
  const after = Number(afterRow) || 0;

  // Para la carga inicial se lee la ventana completa; para sondeos, solo las filas nuevas
  const data = readMessages_(after > 0 ? after + 1 : 2);
  const incoming = {};   // email -> última fila que me han enviado
  const mensajes = [];

  // Para los avisos de "no leído" se necesita ver también las filas ya conocidas,
  // por eso en sondeos se amplía con una lectura corta de las últimas filas.
  const scan = after > 0 ? readMessages_(Math.max(2, data.lastRow - 100)) : data;
  scan.values.forEach((row, i) => {
    if (norm_(row[4]) === me.email) incoming[norm_(row[2])] = scan.start + i;
  });

  data.values.forEach((row, i) => {
    const n = data.start + i;
    if (n <= after || !contact) return;
    const s = norm_(row[2]), r = norm_(row[4]);
    if ((s === me.email && r === contact) || (s === contact && r === me.email)) {
      mensajes.push(toMessage_(row, n));
    }
  });

  return { contactos: contactos, mensajes: mensajes, incoming: incoming, ultimaFila: data.lastRow };
}

function getConversation(contactEmail, afterRow) {
  return sync(contactEmail, afterRow).mensajes;
}

function heartbeat() {
  const me = getIdentity_();
  touchUser_(me);
  return getContacts_(me);
}

function sendMessage(contactEmail, message) {
  const me = getIdentity_();
  const to = norm_(contactEmail);
  const text = String(message || '').trim();
  if (!to) throw new Error('Selecciona un contacto.');
  if (to === me.email) throw new Error('No puedes enviarte mensajes a ti mismo.');
  if (!text) throw new Error('El mensaje está vacío.');
  if (text.length > APP.MAX_LEN) throw new Error('El mensaje supera los ' + APP.MAX_LEN + ' caracteres.');
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
      mensaje: text
    };
    const range = sheet.getRange(rowNumber, 1, 1, 6);
    // Formato de texto ANTES de escribir: evita que "=FORMULA()", "1/2" o "007" se interpreten
    range.setNumberFormats([['@', 'yyyy-mm-dd hh:mm:ss.000', '@', '@', '@', '@']]);
    range.setValues([[item.id, now, item.emisorEmail, item.emisorNombre, item.receptorEmail, item.mensaje]]);
    SpreadsheetApp.flush();
    return item;
  } finally {
    lock.releaseLock();
  }
}
