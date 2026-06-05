function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const ricariche = body.ricariche || [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheets()[0];

    // Intestazioni
    sheet.getRange(1, 1, 1, 10).setValues([[
      'Data','KM','KM Parziali','% Prima','% Dopo','kWh Eff','€/kWh','Costo €','kWh/100km','Luogo'
    ]]);

    // Cancella righe precedenti
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 10).clearContent();

    // Scrivi dati
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

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: 'Usa POST' }))
    .setMimeType(ContentService.MimeType.JSON);
}