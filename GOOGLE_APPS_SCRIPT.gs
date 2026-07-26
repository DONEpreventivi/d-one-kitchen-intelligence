const SECRET_TOKEN = 'CAMBIA-QUESTO-CODICE';
const DB_SHEET = '_D_ONE_DB';

function doGet() {
  return jsonResponse({ok:true, service:'D.ONE Kitchen Sync', version:'0.7'});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.token !== SECRET_TOKEN) return jsonResponse({ok:false,error:'Codice segreto non valido'});
    const sheet = getDbSheet();
    if (body.action === 'pull') return jsonResponse({ok:true,data:readDatabase(sheet)});
    if (body.action !== 'sync') return jsonResponse({ok:false,error:'Azione non riconosciuta'});

    let appliedDays = 0;
    (body.days || []).forEach(day => {
      if (!day || !day.date) return;
      const key = 'DAY:' + day.date;
      const current = readKey(sheet,key);
      const currentTs = current ? Number(current.updatedAt || 0) : 0;
      const incomingTs = Number(day.updatedAt || Date.now());
      if (!current || incomingTs >= currentTs) {
        writeKey(sheet,key,{...day,updatedAt:incomingTs,device:body.device || ''});
        appliedDays++;
      }
    });

    if (Array.isArray(body.menus)) {
      const currentMenus = readKey(sheet,'MENUS');
      const currentTs = currentMenus ? Number(currentMenus.updatedAt || 0) : 0;
      const incomingTs = Number(body.menusUpdatedAt || Date.now());
      if (!currentMenus || incomingTs >= currentTs) writeKey(sheet,'MENUS',{items:body.menus,updatedAt:incomingTs,device:body.device || ''});
    }

    appendLog(body.device || '', appliedDays, (body.days || []).length, !!body.menus);
    return jsonResponse({ok:true,appliedDays:appliedDays,data:readDatabase(sheet)});
  } catch (err) {
    return jsonResponse({ok:false,error:String(err && err.message || err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function getDbSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DB_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DB_SHEET);
    sheet.getRange(1,1,1,4).setValues([['KEY','JSON','UPDATED_AT','DEVICE']]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readDatabase(sheet) {
  const values = sheet.getDataRange().getValues();
  const days = [];
  let menus = null, menusUpdatedAt = 0;
  for (let i=1;i<values.length;i++) {
    const key = String(values[i][0] || '');
    if (!key) continue;
    let value;
    try { value = JSON.parse(values[i][1]); } catch (_) { continue; }
    if (key.indexOf('DAY:') === 0) days.push(value);
    if (key === 'MENUS') { menus = value.items || []; menusUpdatedAt = Number(value.updatedAt || 0); }
  }
  days.sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  return {days:days,menus:menus,menusUpdatedAt:menusUpdatedAt,serverTime:Date.now()};
}

function readKey(sheet,key) {
  const row = findRow(sheet,key);
  if (!row) return null;
  try { return JSON.parse(sheet.getRange(row,2).getValue()); } catch (_) { return null; }
}

function writeKey(sheet,key,value) {
  let row = findRow(sheet,key);
  if (!row) row = sheet.getLastRow()+1;
  sheet.getRange(row,1,1,4).setValues([[key,JSON.stringify(value),new Date(),value.device || '']]);
}

function findRow(sheet,key) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const cell = sheet.getRange(2,1,last-1,1).createTextFinder(key).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function appendLog(device,applied,received,menus) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName('SYNC_LOG');
  if(!sheet){sheet=ss.insertSheet('SYNC_LOG');sheet.appendRow(['TIMESTAMP','DEVICE','GIORNATE_APPLICATE','GIORNATE_RICEVUTE','MENU_INVIATO']);sheet.setFrozenRows(1)}
  sheet.appendRow([new Date(),device,applied,received,menus?'SI':'NO']);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
