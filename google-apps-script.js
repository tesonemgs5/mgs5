// ============================================
// GOOGLE APPS SCRIPT — MGS5 Tracker Sync
// Incolla questo codice in:
// Google Sheet → Estensioni → Apps Script
// ============================================

const SHEET_NAME = 'Foglio1'; // cambia se il tuo foglio ha un nome diverso

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (data.action === 'sync') {
      // Cancella tutto tranne header (riga 1)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

      // Scrivi header se foglio vuoto
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Data','KM Totali','KM Parziali','% Prima','% Dopo','kWh Effettivi','Prezzo €/kWh','Costo €','kWh/100km','Luogo']);
      }

      // Scrivi tutti i dati
      data.ricariche.forEach(r => {
        sheet.appendRow([
          r.data,
          r.km || '',
          r.kmParziali || '',
          r.pctPrima,
          r.pctDopo,
          r.kwhEff,
          r.prezzoKwh,
          r.costo,
          r.kwh100 || '',
          r.luogo || ''
        ]);
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  const rows = sheet.getDataRange().getValues();
  const ricariche = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r[0]) continue;
    ricariche.push({
      data: r[0], km: r[1]||null, kmParziali: r[2]||null,
      pctPrima: r[3], pctDopo: r[4], kwhEff: r[5],
      prezzoKwh: r[6], costo: r[7], kwh100: r[8]||null, luogo: r[9]||null,
      kwhTeor: 0
    });
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ricariche }))
    .setMimeType(ContentService.MimeType.JSON);
}
