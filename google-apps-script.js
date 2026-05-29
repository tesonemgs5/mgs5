// ============================================
// GOOGLE APPS SCRIPT — MGS5 Tracker Sync
// Incolla questo codice in:
// Google Sheet → Estensioni → Apps Script
// ============================================



function doPost(e) { return doGet(e); }

function doGet(e) {
  try {
    const action = e.parameter.action || 'read';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];

    if (action === 'write') {
      const ricariche = JSON.parse(decodeURIComponent(e.parameter.data || '[]'));
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
      ricariche.forEach((r, i) => {
        sheet.getRange(i + 2, 1, 1, 10).setValues([[
          r.data, r.km || '', r.kmParziali || '',
          r.pctPrima, r.pctDopo, r.kwhEff,
          r.prezzoKwh, r.costo, r.kwh100 || '', r.luogo || ''
        ]]);
      });
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const rows = sheet.getDataRange().getValues();
    const data = [];
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      if (!r[0]) continue;
      data.push({
        data: r[0], km: r[1]||null, kmParziali: r[2]||null,
        pctPrima: r[3], pctDopo: r[4], kwhEff: r[5],
        prezzoKwh: r[6], costo: r[7], kwh100: r[8]||null, luogo: r[9]||null,
        kwhTeor: 0
      });
    }
    return ContentService
      .createTextOutput(JSON.stringify({ data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
