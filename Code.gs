/**
 * RetroTalk 2000 - Chat tipo Messenger sobre Google Sheets
 *
 * DESPLIEGUE: Implementar > Aplicación web
 *   - Ejecutar como: "Usuario que accede a la aplicación web"
 *   - Quién tiene acceso: "Cualquier usuario de tu dominio" (Workspace)
 * El HTML debe llamarse exactamente "Index" (Index.html).
 *
 * ENTREGA INSTANTÁNEA: cada mensaje se deposita primero en un "buzón" en caché del receptor
 * (la petición en espera del receptor lo recoge en ~150 ms) y después se guarda en la hoja.
 * La hoja es el historial permanente; la caché es el canal rápido.
 */
const APP = {
  USERS_SHEET: 'Usuarios',
  MESSAGES_SHEET: 'Mensajes',
  ONLINE_MS: 45000,     // sin actividad durante este tiempo => desconectado
  TOUCH_MS: 15000,      // como máximo se escribe la actividad cada 15 s por usuario
  LONGPOLL_MS: 15000,   // tiempo máximo que el servidor mantiene una petición esperando novedades
  POLL_SLEEP_MS: 150,   // cada cuánto mira la caché mientras espera (menor = más inmediato)
  MAILBOX_MS: 1800000,  // los mensajes del buzón rápido se conservan 30 min
  MAX_ROWS: 500,        // filas recientes de Mensajes que se leen en cada consulta
  MAX_LEN: 1000,
  NUDGE_MS: 10000,      // tiempo mínimo entre zumbidos del mismo usuario
  STATES: ['disponible', 'llamada', 'cafe', 'desconectado'],
  SPREADSHEET_ID: ''    // Vacío si el script está vinculado a la hoja; si no, pega aquí el ID.
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

function bumpPresence_() {
  CacheService.getScriptCache().put('presence', String(Date.now()), 21600);
}

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

/** Migra hojas creadas con versiones anteriores (columnas "estado" y "tipo"). */
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

/** ¿Existe el usuario? Usa caché para no leer la hoja en cada envío. */
function knownUser_(email) {
  const cache = CacheService.getScriptCache();
  try {
    const raw = cache.get('known_users');
    if (raw && JSON.parse(raw).indexOf(email) >= 0) return true;
  } catch (_) {}
  const sheet = sheet_(APP.USERS_SHEET);
  const last = sheet.getLastRow();
  const emails = last >= 2 ? sheet.getRange(2, 1, last - 1, 1).getValues().map(r => norm_(r[0])) : [];
  cache.put('known_users', JSON.stringify(emails), 300);
  return emails.indexOf(email) >= 0;
}

/** Registra actividad (limitado por caché). Avisa a los demás si el usuario pasa de desconectado a conectado. */
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
      cache.remove('known_users');
      bumpPresence_();
    } else {
      const prev = new Date(sheet.getRange(row, 3).getValue()).getTime();
      sheet.getRange(row, 2, 1, 2).setValues([[user.nombre, new Date()]]);
      if (!(prev >= Date.now() - APP.ONLINE_MS)) bumpPresence_();
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
      const estado = active ? saved : 'desconectado';
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

/* ---------- Buzón rápido en caché ---------- */

function mailboxRead_(cache, email) {
  try { return JSON.parse(cache.get('mbox_' + email) || '[]'); } catch (_) { return []; }
}

/** Deposita el mensaje en el buzón del receptor y "despierta" su petición en espera. */
function pushMailbox_(to, item) {
  const cache = CacheService.getScriptCache();
  for (let attempt = 0; attempt < 3; attempt++) {
    const cutoff = Date.now() - APP.MAILBOX_MS;
    let arr = mailboxRead_(cache, to).filter(m => m.id !== item.id && new Date(m.fecha).getTime() > cutoff);
    const prev = arr.length ? Number(arr[arr.length - 1].seq) : 0;
    item.seq = Math.max(Date.now() * 1000 + Math.floor(Math.random() * 1000), prev + 1);
    arr.push(item);
    let json = JSON.stringify(arr);
    while ((arr.length > 40 || json.length > 30000) && arr.length > 1) { arr.shift(); json = JSON.stringify(arr); }
    cache.put('mbox_' + to, json, 21600);
    cache.put('seq_' + to, String(item.seq), 21600);
    // Verificación (sin bloqueo): si otro envío simultáneo pisó el buzón, se reintenta
    if (mailboxRead_(cache, to).some(m => m.id === item.id)) return;
  }
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
 * Sincronización con long-polling + buzón rápido.
 * - Camino rápido: si mientras esperaba llegó un mensaje, responde SOLO con él (sin leer la hoja).
 * - Camino completo: carga inicial, cambios de presencia o fin del tiempo de espera
 *   (contactos + historial desde la hoja; actúa además como red de seguridad).
 *
 * @param {string} contactEmail contacto abierto ('' si ninguno)
 * @param {number} afterRow     última fila de la hoja ya recibida para ese chat
 * @param {Object} opts { wait, seq, itemSeq, pres, passive, loop, first }
 */
function sync(contactEmail, afterRow, opts) {
  opts = opts || {};
  const me = getIdentity_();
  const cache = CacheService.getScriptCache();
  const kSeq = 'seq_' + me.email, kTouch = 'touch_' + me.email, kLoop = 'loop_' + me.email;
  const passive = !!opts.passive;
  const loop = String(opts.loop || '');
  const clientSeq = String(opts.seq || '0');
  const clientPres = String(opts.pres || '');
  const itemSeq = Number(opts.itemSeq) || 0;
  const waitMs = Math.min(Math.max(Number(opts.wait) || 0, 0), APP.LONGPOLL_MS);

  if (loop && opts.first) cache.put(kLoop, loop, 21600);   // este bucle sustituye a los anteriores
  if (!passive) touchUser_(me);

  const snapshot = () => {
    const g = cache.getAll([kSeq, 'presence', kTouch, kLoop]);
    return {
      seq: g[kSeq] || '0',
      pres: g.presence || '',
      touch: !!g[kTouch],
      stale: !!loop && !!g[kLoop] && g[kLoop] !== loop
    };
  };
  const changed = s => s.seq !== clientSeq || s.pres !== clientPres;

  let st = snapshot();
  if (waitMs > 0 && !changed(st) && !st.stale) {
    const end = Date.now() + waitMs;
    while (Date.now() < end) {
      Utilities.sleep(APP.POLL_SLEEP_MS);
      st = snapshot();
      if (st.stale || changed(st)) break;
      if (!passive && !st.touch) touchUser_(me);
    }
  }

  // Mensajes nuevos del buzón rápido
  let items = [], maxSeq = itemSeq;
  if (st.seq !== clientSeq || itemSeq === 0) {
    const all = mailboxRead_(cache, me.email);
    all.forEach(m => { if (Number(m.seq) > maxSeq) maxSeq = Number(m.seq); });
    items = all.filter(m => Number(m.seq) > itemSeq);
  }

  // Camino rápido: solo llegaron mensajes
  if (waitMs > 0 && items.length && st.pres === clientPres && !st.stale) {
    return { full: false, items: items, seq: st.seq, itemSeq: maxSeq, pres: st.pres, stale: false };
  }

  // Camino completo
  const contactos = getContacts_(me);
  const contact = norm_(contactEmail);
  const after = Number(afterRow) || 0;
  const mensajes = [];
  if (contact) {
    const data = readMessages_(after > 0 ? after + 1 : 2);
    data.values.forEach((row, i) => {
      const n = data.start + i;
      if (n <= after) return;
      const s = norm_(row[2]), r = norm_(row[4]);
      if ((s === me.email && r === contact) || (s === contact && r === me.email)) {
        mensajes.push(toMessage_(row, n));
      }
    });
  }
  return {
    full: true, contactos: contactos, mensajes: mensajes, items: items,
    seq: st.seq, itemSeq: maxSeq, pres: st.pres, stale: st.stale
  };
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
  bumpPresence_();
  return estado;
}

/** El cliente la llama al quedar ausente o al cerrar la página. */
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
  bumpPresence_();
}

/** Guarda el mensaje en la hoja (historial permanente). */
function persist_(item) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = sheet_(APP.MESSAGES_SHEET);
    const row = sheet.getLastRow() + 1;
    const range = sheet.getRange(row, 1, 1, 7);
    range.setNumberFormats([['@', 'yyyy-mm-dd hh:mm:ss.000', '@', '@', '@', '@', '@']]);
    range.setValues([[item.id, new Date(item.fecha), item.emisorEmail, item.emisorNombre, item.receptorEmail, item.mensaje, item.tipo]]);
    item.fila = row;
  } finally {
    lock.releaseLock();
  }
}

function append_(me, to, text, tipo) {
  if (!to) throw new Error('Selecciona un contacto.');
  if (to === me.email) throw new Error('No puedes enviarte mensajes a ti mismo.');
  if (!knownUser_(to)) throw new Error('El contacto no existe.');
  touchUser_(me);
  const item = {
    fila: 0,
    id: Utilities.getUuid(),
    fecha: new Date().toISOString(),
    emisorEmail: me.email,
    emisorNombre: me.nombre,
    receptorEmail: to,
    mensaje: text,
    tipo: tipo
  };
  pushMailbox_(to, item);   // 1) entrega inmediata al receptor
  persist_(item);           // 2) guardado en la hoja
  return item;
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
  cache.put(key, '1', Math.max(1, Math.ceil(APP.NUDGE_MS / 1000)));
  return append_(me, norm_(contactEmail), '¡Zumbido!', 'zumbido');
}
