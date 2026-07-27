/**
 * D.ONE Kitchen Intelligence — Google Sheets backend v0.9.6.1
 *
 * IMPORTANTE:
 * Dopo aver sostituito il vecchio script:
 * 1. Salva.
 * 2. Distribuisci > Gestisci deployment.
 * 3. Modifica (matita) > Nuova versione > Esegui il deployment.
 * L'URL /exec rimane lo stesso.
 */

const SHARED_SECRET = 'CAMBIA-QUESTO-CODICE-CON-LO-STESSO-GIA-USATO';

const SHEETS = {
  DAYS: 'GIORNATE',
  SALES: 'VENDITE',
  MENUS: 'MENU',
  LOG: 'SYNC_LOG'
};

const MESSAGE_TYPE = 'DONE_KITCHEN_SYNC_V09';

function doGet() {
  ensureDatabase_();
  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      service: 'D.ONE Kitchen Sync',
      version: '0.9.6',
      status: 'online',
      time: new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const requestId = String((e && e.parameter && e.parameter.requestId) || '');
  const iframeTransport = String((e && e.parameter && e.parameter.transport) || '') === 'iframe';

  let result;
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);

    let request = {};
    const formPayload = e && e.parameter && e.parameter.payload;

    if (formPayload) {
      request = JSON.parse(formPayload);
    } else {
      request = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    }

    if (!request.token || request.token !== SHARED_SECRET) {
      result = { ok: false, error: 'Codice segreto non valido' };
    } else {
      ensureDatabase_();

      if (request.action === 'ping') {
        result = {
          ok: true,
          status: 'online',
          version: '0.9.6',
          time: new Date().toISOString()
        };
      } else if (request.action === 'pull') {
        result = {
          ok: true,
          data: readDatabase_()
        };
      } else if (request.action === 'sync') {
        const device = String(request.device || 'Dispositivo senza nome');
        const days = Array.isArray(request.days) ? request.days : [];
        const appliedDays = bulkUpsertDays_(days, device);

        if (Array.isArray(request.menus)) {
          replaceMenus_(
            request.menus,
            Number(request.menusUpdatedAt || Date.now()),
            device
          );
        }

        appendLog_(device, 'sync', appliedDays, days.length);

        result = {
          ok: true,
          appliedDays: appliedDays,
          version: '0.9.6',
          time: new Date().toISOString()
        };
      } else {
        result = { ok: false, error: 'Azione non riconosciuta' };
      }
    }
  } catch (error) {
    result = {
      ok: false,
      error: error && error.message ? error.message : String(error)
    };
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }

  if (iframeTransport) {
    return iframeResponse_(requestId || String((e.parameter && e.parameter.requestId) || ''), result);
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function iframeResponse_(requestId, result) {
  const safeRequestId = JSON.stringify(String(requestId || ''));
  const safeResult = JSON.stringify(result).replace(/</g, '\\u003c');

  const html = [
    '<!doctype html><html><head><meta charset="utf-8"></head><body>',
    '<script>',
    '(function(){',
    'var message={type:', JSON.stringify(MESSAGE_TYPE),
    ',requestId:', safeRequestId,
    ',result:', safeResult, '};',
    'try{window.parent.postMessage(message,"*");}catch(e){}try{window.top.postMessage(message,"*");}catch(e){}',
    '})();',
    '<\/script>',
    '</body></html>'
  ].join('');

  return HtmlService
    .createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function ensureDatabase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ensureSheet_(ss, SHEETS.DAYS, [
    'data','aggiornato_il','coperti','incasso_locale','note','dispositivo'
  ]);

  ensureSheet_(ss, SHEETS.SALES, [
    'data','id_prodotto','prodotto','categoria','quantita',
    'prezzo_unitario','totale','aggiornato_il','dispositivo'
  ]);

  ensureSheet_(ss, SHEETS.MENUS, [
    'ordine','id_menu','nome_menu','inizio','fine',
    'aggiornato_il','dispositivo','json_menu'
  ]);

  ensureSheet_(ss, SHEETS.LOG, [
    'data_ora','dispositivo','azione','giornate_applicate','giornate_inviate'
  ]);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}


function dedupeRows_(rows) {
  const seen = {};
  const clean = [];
  (rows || []).forEach(function(row) {
    const key = String(row.productId || '');
    if (seen[key]) return;
    seen[key] = true;
    clean.push(row);
  });
  return clean;
}

function bulkUpsertDays_(incomingDays, device) {
  incomingDays = Array.isArray(incomingDays) ? incomingDays : [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daysSheet = ss.getSheetByName(SHEETS.DAYS);
  const salesSheet = ss.getSheetByName(SHEETS.SALES);

  const existing = {};
  readDays_().forEach(function(day) {
    existing[day.date] = day;
  });

  let applied = 0;
  incomingDays.forEach(function(day) {
    if (!day || !day.date) return;
    const date = String(day.date);
    const incomingTs = Number(day.updatedAt || Date.now());

    // Una giornata inviata esplicitamente dal dispositivo è autorevole:
    // sostituisce sempre la stessa data presente nel database.
    existing[date] = {
      date: date,
      updatedAt: incomingTs,
      covers: Number(day.covers || 0),
      venueRevenue: Number(day.venueRevenue || 0),
      notes: String(day.notes || ''),
      rows: dedupeRows_(Array.isArray(day.rows) ? day.rows : []),
      device: device
    };
    applied++;
  });

  const allDays = Object.keys(existing).sort().map(function(date) {
    return existing[date];
  });

  if (daysSheet.getLastRow() > 1) {
    daysSheet.getRange(2, 1, daysSheet.getLastRow() - 1, daysSheet.getLastColumn()).clearContent();
  }
  if (salesSheet.getLastRow() > 1) {
    salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, salesSheet.getLastColumn()).clearContent();
  }

  if (allDays.length) {
    const dayValues = allDays.map(function(day) {
      return [
        day.date,
        Number(day.updatedAt || 0),
        Number(day.covers || 0),
        Number(day.venueRevenue || 0),
        String(day.notes || ''),
        String(day.device || device || '')
      ];
    });
    daysSheet.getRange(2, 1, dayValues.length, 6).setValues(dayValues);
  }

  const saleValues = [];
  allDays.forEach(function(day) {
    (day.rows || []).forEach(function(row) {
      const qty = Number(row.qty || 0);
      const price = Number(row.price || 0);
      saleValues.push([
        day.date,
        String(row.productId || ''),
        String(row.name || ''),
        String(row.category || ''),
        qty,
        price,
        qty * price,
        Number(day.updatedAt || 0),
        String(day.device || device || '')
      ]);
    });
  });

  if (saleValues.length) {
    salesSheet.getRange(2, 1, saleValues.length, 9).setValues(saleValues);
  }

  sortDataSheets_();
  SpreadsheetApp.flush();
  return applied;
}

function upsertDay_(day, device) {
  if (!day || !day.date) return false;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daysSheet = ss.getSheetByName(SHEETS.DAYS);
  const salesSheet = ss.getSheetByName(SHEETS.SALES);

  const date = String(day.date);
  const incomingUpdatedAt = Number(day.updatedAt || Date.now());
  const existingRow = findDateRow_(daysSheet, date);

  if (existingRow) {
    const existingUpdatedAt = Number(daysSheet.getRange(existingRow, 2).getValue() || 0);
    if (existingUpdatedAt > incomingUpdatedAt) return false;

    daysSheet.getRange(existingRow, 1, 1, 6).setValues([[
      date,
      incomingUpdatedAt,
      Number(day.covers || 0),
      Number(day.venueRevenue || 0),
      String(day.notes || ''),
      device
    ]]);
  } else {
    daysSheet.appendRow([
      date,
      incomingUpdatedAt,
      Number(day.covers || 0),
      Number(day.venueRevenue || 0),
      String(day.notes || ''),
      device
    ]);
  }

  deleteSalesForDate_(salesSheet, date);

  const rows = Array.isArray(day.rows) ? day.rows : [];
  if (rows.length) {
    const values = rows.map(function(row) {
      const qty = Number(row.qty || 0);
      const price = Number(row.price || 0);
      return [
        date,
        String(row.productId || ''),
        String(row.name || ''),
        String(row.category || ''),
        qty,
        price,
        qty * price,
        incomingUpdatedAt,
        device
      ];
    });

    salesSheet.getRange(salesSheet.getLastRow() + 1, 1, values.length, 9).setValues(values);
  }

  sortDataSheets_();
  return true;
}

function findDateRow_(sheet, date) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === date) return i + 2;
  }
  return 0;
}

function deleteSalesForDate_(sheet, date) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  const rowsToDelete = [];

  values.forEach(function(row, index) {
    if (String(row[0]) === date) rowsToDelete.push(index + 2);
  });

  rowsToDelete.reverse().forEach(function(rowNumber) {
    sheet.deleteRow(rowNumber);
  });
}

function replaceMenus_(menus, updatedAt, device) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MENUS);

  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  if (!menus.length) return;

  const values = menus.map(function(menu, index) {
    return [
      index + 1,
      String(menu.id || ('menu-' + (index + 1))),
      String(menu.name || ''),
      String(menu.start || ''),
      String(menu.end || ''),
      Number(updatedAt || Date.now()),
      device,
      JSON.stringify(menu)
    ];
  });

  sheet.getRange(2, 1, values.length, 8).setValues(values);
}

function readDatabase_() {
  return {
    days: readDays_(),
    menus: readMenus_(),
    menusUpdatedAt: readMenusUpdatedAt_()
  };
}

function readDays_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daysSheet = ss.getSheetByName(SHEETS.DAYS);
  const salesSheet = ss.getSheetByName(SHEETS.SALES);
  const salesByDate = {};

  if (salesSheet.getLastRow() > 1) {
    const sales = salesSheet.getRange(2, 1, salesSheet.getLastRow() - 1, 9).getValues();

    const seenSales = {};
    sales.forEach(function(row) {
      const date = formatDateValue_(row[0]);
      const productId = String(row[1] || '');
      const key = date + '__' + productId;
      if (seenSales[key]) return;
      seenSales[key] = true;
      if (!salesByDate[date]) salesByDate[date] = [];
      salesByDate[date].push({
        date: date,
        productId: productId,
        name: String(row[2] || ''),
        category: String(row[3] || ''),
        qty: Number(row[4] || 0),
        price: Number(row[5] || 0)
      });
    });
  }

  if (daysSheet.getLastRow() < 2) return [];

  return daysSheet.getRange(2, 1, daysSheet.getLastRow() - 1, 6).getValues()
    .map(function(row) {
      const date = formatDateValue_(row[0]);
      return {
        date: date,
        updatedAt: Number(row[1] || 0),
        covers: Number(row[2] || 0),
        venueRevenue: Number(row[3] || 0),
        notes: String(row[4] || ''),
        rows: salesByDate[date] || []
      };
    })
    .sort(function(a, b) { return a.date.localeCompare(b.date); });
}

function readMenus_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MENUS);
  if (sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues()
    .sort(function(a, b) { return Number(a[0] || 0) - Number(b[0] || 0); })
    .map(function(row) {
      try { return JSON.parse(String(row[7] || '{}')); }
      catch (_) { return null; }
    })
    .filter(Boolean);
}

function readMenusUpdatedAt_() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MENUS);
  if (sheet.getLastRow() < 2) return 0;

  const values = sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getValues().flat().map(Number);
  return values.length ? Math.max.apply(null, values) : 0;
}

function sortDataSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const days = ss.getSheetByName(SHEETS.DAYS);
  const sales = ss.getSheetByName(SHEETS.SALES);

  if (days.getLastRow() > 2) {
    days.getRange(2, 1, days.getLastRow() - 1, days.getLastColumn())
      .sort([{ column: 1, ascending: true }]);
  }

  if (sales.getLastRow() > 2) {
    sales.getRange(2, 1, sales.getLastRow() - 1, sales.getLastColumn())
      .sort([
        { column: 1, ascending: true },
        { column: 4, ascending: true },
        { column: 3, ascending: true }
      ]);
  }
}

function appendLog_(device, action, appliedDays, sentDays) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.LOG).appendRow([
    new Date(),
    device,
    action,
    Number(appliedDays || 0),
    Number(sentDays || 0)
  ]);
}

function formatDateValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value || '');
}
