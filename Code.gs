const CONFIG = {
  SHEET_NAME: 'Mensajes',
  USERS_SHEET: 'Usuarios',
  MAX_MESSAGES: 120,
  POLL_LIMIT: 80
};

function doGet() {
  setupSheets_();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('RetroTalk 2000')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.SHEET_NAME);
    sh.appendRow(['id','timestamp','email','displayName','text']);
    sh.setFrozenRows(1);
  }
  let users = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!users) {
    users = ss.insertSheet(CONFIG.USERS_SHEET);
    users.appendRow(['email','displayName','lastSeen','status']);
    users.setFrozenRows(1);
  }
}

function getBootstrap() {
  setupSheets_();
  const email = Session.getActiveUser().getEmail() || 'usuario@interno.local';
  const name = email.split('@')[0].replace(/[._-]+/g,' ').replace(/\b\w/g, c => c.toUpperCase());
  touchUser_(email, name);
  return {
    me: { email, name },
    messages: getMessages(''),
    users: getOnlineUsers()
  };
}

function sendMessage(text) {
  text = String(text || '').trim();
  if (!text) throw new Error('El mensaje está vacío.');
  if (text.length > 1000) throw new Error('Máximo 1000 caracteres.');

  setupSheets_();
  const email = Session.getActiveUser().getEmail() || 'usuario@interno.local';
  const name = email.split('@')[0].replace(/[._-]+/g,' ').replace(/\b\w/g, c => c.toUpperCase());
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
    const id = Utilities.getUuid();
    const now = new Date();
    sh.appendRow([id, now, email, name, text]);
    touchUser_(email, name);
    return {id, timestamp: now.toISOString(), email, displayName: name, text};
  } finally {
    lock.releaseLock();
  }
}

function getMessages(afterIso) {
  setupSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.SHEET_NAME);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const start = Math.max(2, last - CONFIG.MAX_MESSAGES + 1);
  const rows = sh.getRange(start, 1, last - start + 1, 5).getValues();
  const after = afterIso ? new Date(afterIso).getTime() : 0;
  return rows.map(r => ({
    id: r[0], timestamp: new Date(r[1]).toISOString(), email: r[2], displayName: r[3], text: r[4]
  })).filter(m => new Date(m.timestamp).getTime() > after).slice(-CONFIG.POLL_LIMIT);
}

function heartbeat() {
  const email = Session.getActiveUser().getEmail() || 'usuario@interno.local';
  const name = email.split('@')[0].replace(/[._-]+/g,' ').replace(/\b\w/g, c => c.toUpperCase());
  touchUser_(email, name);
  return getOnlineUsers();
}

function touchUser_(email, name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const values = sh.getDataRange().getValues();
  const now = new Date();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === email) {
      sh.getRange(i + 1, 2, 1, 3).setValues([[name, now, 'online']]);
      return;
    }
  }
  sh.appendRow([email, name, now, 'online']);
}

function getOnlineUsers() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CONFIG.USERS_SHEET);
  const values = sh.getDataRange().getValues().slice(1);
  const cutoff = Date.now() - 45000;
  return values.map(r => ({
    email:r[0], name:r[1], lastSeen:new Date(r[2]).toISOString(), online:new Date(r[2]).getTime() >= cutoff
  })).sort((a,b) => Number(b.online)-Number(a.online) || a.name.localeCompare(b.name));
}
