function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(name);
  if (!s) s = ss.insertSheet(name);
  return s;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.action === 'new') {
      const s = getSheet('RDV');
      s.appendRow([
        new Date().toLocaleString('fr-FR'),
        data.prenom || '',
        data.nom || '',
        data.telephone || '',
        data.prestation || '',
        data.date || '',
        data.creneau || '',
        'En attente',
        data.message || ''
      ]);
      notifyTelegramNewRdv(data);
      return jsonResponse({ success: true });
    }

    if (data.action === 'update') {
      const s = getSheet('RDV');
      s.getRange(parseInt(data.row), 8).setValue(data.status);
      return jsonResponse({ success: true });
    }

    if (data.action === 'save_disponibilites') {
      const s = getSheet('Config');
      s.getRange('A1').setValue(JSON.stringify(data.disponibilites));
      return jsonResponse({ success: true });
    }

    if (data.action === 'save_tarifs') {
      const s = getSheet('Config');
      s.getRange('B1').setValue(JSON.stringify(data.tarifs));
      return jsonResponse({ success: true });
    }

    return jsonResponse({ success: false, error: 'Action inconnue' });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function doGet(e) {
  try {
    const type = (e.parameter && e.parameter.type) || 'all';
    const config = getSheet('Config');
    const dispoVal = config.getRange('A1').getValue();
    const tarifsVal = config.getRange('B1').getValue();

    if (type === 'disponibilites') {
      return jsonResponse({ success: true, disponibilites: dispoVal ? JSON.parse(dispoVal) : {} });
    }

    if (type === 'tarifs') {
      return jsonResponse({ success: true, tarifs: tarifsVal ? JSON.parse(tarifsVal) : null });
    }

    const rdvSheet = getSheet('RDV');
    const rows = rdvSheet.getDataRange().getValues();
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
        status: rows[i][7] || 'En attente',
        message: rows[i][8] || ''
      });
    }

    return jsonResponse({
      success: true,
      bookings: bookings,
      disponibilites: dispoVal ? JSON.parse(dispoVal) : {},
      tarifs: tarifsVal ? JSON.parse(tarifsVal) : null
    });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function notifyTelegramNewRdv(data) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) return; // Notifications non configurées

  const lines = [
    '📅 Nouvelle demande de RDV',
    `${data.prenom || ''} ${data.nom || ''}`.trim(),
    data.prestation ? `Prestation : ${data.prestation}` : '',
    data.date ? `Date : ${data.date}` : '',
    data.creneau ? `Créneau : ${data.creneau}` : '',
    data.telephone ? `Téléphone : ${data.telephone}` : '',
    data.message ? `Message : ${data.message}` : ''
  ].filter(Boolean);

  try {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: lines.join('\n') }),
      muteHttpExceptions: true
    });
  } catch (err) {
    Logger.log('Erreur notification Telegram: ' + err);
  }
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function setup() {
  const rdv = getSheet('RDV');
  if (rdv.getLastRow() === 0) {
    rdv.getRange(1, 1, 1, 8).setValues([['Date soumission', 'Prénom', 'Nom', 'Téléphone', 'Prestation', 'Date RDV', 'Créneau', 'Statut']]);
    rdv.getRange(1, 1, 1, 8).setFontWeight('bold');
    rdv.setFrozenRows(1);
  }
  getSheet('Config');
  Logger.log('Setup terminé !');
}
