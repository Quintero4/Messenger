const APP = {
  USERS_SHEET: 'Usuarios',
  MESSAGES_SHEET: 'Mensajes',
  ONLINE_MS: 45000,
  MAX_ROWS: 500
};

function doGet() {
  setupDatabase_();
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('RetroTalk 2000');
}

function setupDatabase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let users = ss.getSheetByName(APP.USERS_SHEET);
  if (!users) {
    users = ss.insertSheet(APP.USERS_SHEET);
    users.appendRow(['email', 'nombre', 'ultimaActividad']);
    users.setFrozenRows(1);
  }
  let messages = ss.getSheetByName(APP.MESSAGES_SHEET);
  if (!messages) {
    messages = ss.insertSheet(APP.MESSAGES_SHEET);
    messages.appendRow(['id', 'fecha', 'emisorEmail', 'emisorNombre', 'receptorEmail', 'mensaje']);
    messages.setFrozenRows(1);
  }
}

function getIdentity_() {
  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('No se pudo identificar al usuario. Publica la Web App para usuarios de tu dominio y ejecútala como usuario que accede.');
  }
  const nombre = email.split('@')[0]
    .replace(/[._-]+/g, ' ')
    .replace(/\b\w/g, c => c.toUpperCase());
  return { email, nombre };
}

function touchUser_(user) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(APP.USERS_SHEET);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === user.email) {
      sheet.getRange(i + 1, 2, 1, 2).setValues([[user.nombre, new Date()]]);
      return;
    }
  }
  sheet.appendRow([user.email, user.nombre, new Date()]);
}

function getContacts_() {
  const me = getIdentity_();
  const sheet = SpreadsheetApp.getActive().getSheetByName(APP.USERS_SHEET);
  const cutoff = Date.now() - APP.ONLINE_MS;
  return sheet.getDataRange().getValues().slice(1)
    .filter(row => row[0] && row[0] !== me.email)
    .map(row => ({
      email: String(row[0]),
      nombre: String(row[1] || row[0]),
      online: new Date(row[2]).getTime() >= cutoff
    }))
    .sort((a, b) => Number(b.online) - Number(a.online) || a.nombre.localeCompare(b.nombre));
}

function bootstrap() {
  setupDatabase_();
  const me = getIdentity_();
  touchUser_(me);
  return { me, contactos: getContacts_() };
}

function heartbeat() {
  const me = getIdentity_();
  touchUser_(me);
  return getContacts_();
}

function getConversation(contactEmail, afterIso) {
  if (!contactEmail) return [];
  const me = getIdentity_();
  touchUser_(me);
  const sheet = SpreadsheetApp.getActive().getSheetByName(APP.MESSAGES_SHEET);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const startRow = Math.max(2, lastRow - APP.MAX_ROWS + 1);
  const rows = sheet.getRange(startRow, 1, lastRow - startRow + 1, 6).getValues();
  const after = afterIso ? new Date(afterIso).getTime() : 0;
  return rows
    .filter(row => {
      const pair = (row[2] === me.email && row[4] === contactEmail) ||
                   (row[2] === contactEmail && row[4] === me.email);
      return pair && new Date(row[1]).getTime() > after;
    })
    .map(row => ({
      id: String(row[0]),
      fecha: new Date(row[1]).toISOString(),
      emisorEmail: String(row[2]),
      emisorNombre: String(row[3]),
      receptorEmail: String(row[4]),
      mensaje: String(row[5])
    }));
}

function sendMessage(contactEmail, message) {
  const me = getIdentity_();
  const text = String(message || '').trim();
  if (!contactEmail) throw new Error('Selecciona un contacto.');
  if (!text) throw new Error('El mensaje está vacío.');
  if (text.length > 1000) throw new Error('El mensaje supera los 1000 caracteres.');
  touchUser_(me);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const now = new Date();
    const item = {
      id: Utilities.getUuid(),
      fecha: now.toISOString(),
      emisorEmail: me.email,
      emisorNombre: me.nombre,
      receptorEmail: contactEmail,
      mensaje: text
    };
    SpreadsheetApp.getActive().getSheetByName(APP.MESSAGES_SHEET)
      .appendRow([item.id, now, item.emisorEmail, item.emisorNombre, item.receptorEmail, item.mensaje]);
    return item;
  } finally {
    lock.releaseLock();
  }
}
