function doPost(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('rdv') || ss.getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'update') {
      sheet.getRange(parseInt(data.row), 7).setValue(data.status);
      return jsonResponse({ success: true });
    }

    sheet.appendRow([
      new Date().toLocaleString('fr-FR'),
      data.prenom || '',
      data.nom || '',
      data.telephone || '',
      data.prestation || '',
      data.date || '',
      data.creneau || '',
      'En attente'
    ]);

    return jsonResponse({ success: true });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('rdv') || ss.getActiveSheet();
    const rows = sheet.getDataRange().getValues();
    const bookings = [];
    for (let i = 1; i < rows.length; i++) {
      bookings.push({
        row: i + 1,
        date_soumission: rows[i][0],
        prenom: rows[i][1],
        nom: rows[i][2],
        telephone: rows[i][3],
        prestation: rows[i][4],
        date: rows[i][5],
        creneau: rows[i][6],
        status: rows[i][7] || 'En attente'
      });
    }
    return jsonResponse({ success: true, bookings: bookings });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, 1, 1, 8).setValues([[
    'Date soumission', 'Prénom', 'Nom', 'Téléphone', 'Prestation', 'Date RDV', 'Créneau', 'Statut'
  ]]);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
