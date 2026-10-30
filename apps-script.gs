// Backend de votos del viaje. Pégalo en Extensiones > Apps Script de una Hoja de Google
// e impleméntalo como aplicación web (Ejecutar como: yo · Acceso: cualquier usuario).

const SHEET = "votos";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(["clave", "nombre", "¿viene?", "presupuesto", "sale de", "actualizado", "json"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  const rows = sheet_().getDataRange().getValues().slice(1);
  const people = {};
  rows.forEach(r => { try { people[r[0]] = JSON.parse(r[6]); } catch (e) {} });
  return json_({ people });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents);
    const key = String(body.key || "").slice(0, 60);
    const entry = body.entry || {};
    if (!key || !entry.name) return json_({ ok: false, error: "Falta el nombre" });
    entry.ts = Date.now();
    const raw = JSON.stringify(entry);
    if (raw.length > 20000) return json_({ ok: false, error: "Respuesta demasiado larga" });

    const sh = sheet_();
    const keys = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().map(r => r[0]);
    const row = [key, entry.name, entry.attend, entry.budget, entry.origin, new Date(), raw];
    const i = keys.indexOf(key);
    if (i > 0) sh.getRange(i + 1, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}
