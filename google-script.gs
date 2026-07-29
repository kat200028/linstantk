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
      const bloque = estDansListeNoire(data.telephone);
      s.appendRow([
        new Date().toLocaleString('fr-FR'),
        data.prenom || '',
        data.nom || '',
        data.telephone || '',
        data.prestation || '',
        data.date || '',
        data.creneau || '',
        bloque ? 'Refusé' : 'En attente',
        data.message || '',
        data.reference || ''
      ]);
      if (!bloque) notifyTelegramNewRdv(data);
      return jsonResponse({ success: true });
    }

    if (data.action === 'admin_add') {
      const s = getSheet('RDV');
      s.appendRow([
        new Date().toLocaleString('fr-FR'),
        data.prenom || '',
        data.nom || '',
        data.telephone || '',
        data.prestation || '',
        data.date || '',
        data.creneau || '',
        data.status || 'Accepté',
        data.message || '',
        data.reference || ''
      ]);
      return jsonResponse({ success: true });
    }

    if (data.action === 'save_blacklist') {
      const s = getSheet('Config');
      s.getRange('C1').setValue(JSON.stringify(data.blacklist || []));
      return jsonResponse({ success: true });
    }

    if (data.action === 'save_client_note') {
      const s = getSheet('Config');
      const val = s.getRange('D1').getValue();
      const notes = val ? JSON.parse(val) : {};
      const tel = normaliserTelServeur(data.telephone);
      if (data.note && data.note.trim()) notes[tel] = data.note;
      else delete notes[tel];
      s.getRange('D1').setValue(JSON.stringify(notes));
      return jsonResponse({ success: true });
    }

    if (data.action === 'login') {
      const hash = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD_HASH');
      if (!hash) return jsonResponse({ success: false, setup: true });
      return jsonResponse({ success: sha256(data.password || '') === hash });
    }

    if (data.action === 'reset_request') {
      const props = PropertiesService.getScriptProperties();
      const email = props.getProperty('ADMIN_EMAIL') || Session.getEffectiveUser().getEmail();
      if (!email) return jsonResponse({ success: false, error: 'Aucune adresse email configurée' });
      const code = String(Math.floor(100000 + Math.random() * 900000));
      props.setProperty('RESET_CODE_HASH', sha256(code));
      props.setProperty('RESET_CODE_EXPIRES', String(Date.now() + 15 * 60 * 1000));
      MailApp.sendEmail(email, "L'Instant K — Réinitialisation du mot de passe admin",
        'Votre code de réinitialisation : ' + code + '\n\nIl est valable 15 minutes.\nSi vous n\'êtes pas à l\'origine de cette demande, ignorez ce message.');
      return jsonResponse({ success: true });
    }

    if (data.action === 'reset_confirm') {
      const props = PropertiesService.getScriptProperties();
      const codeHash = props.getProperty('RESET_CODE_HASH');
      const expires = parseInt(props.getProperty('RESET_CODE_EXPIRES') || '0', 10);
      if (!codeHash || Date.now() > expires) return jsonResponse({ success: false, error: 'Code expiré — redemandez un code' });
      if (sha256(String(data.code || '').trim()) !== codeHash) return jsonResponse({ success: false, error: 'Code incorrect' });
      if (!data.newPassword || String(data.newPassword).length < 6) {
        return jsonResponse({ success: false, error: 'Le nouveau mot de passe doit faire au moins 6 caractères' });
      }
      props.setProperty('ADMIN_PASSWORD_HASH', sha256(String(data.newPassword)));
      props.deleteProperty('RESET_CODE_HASH');
      props.deleteProperty('RESET_CODE_EXPIRES');
      return jsonResponse({ success: true });
    }

    if (data.action === 'change_password') {
      const props = PropertiesService.getScriptProperties();
      const hash = props.getProperty('ADMIN_PASSWORD_HASH');
      if (hash && sha256(data.oldPassword || '') !== hash) {
        return jsonResponse({ success: false, error: 'Ancien mot de passe incorrect' });
      }
      if (!data.newPassword || String(data.newPassword).length < 6) {
        return jsonResponse({ success: false, error: 'Le nouveau mot de passe doit faire au moins 6 caractères' });
      }
      props.setProperty('ADMIN_PASSWORD_HASH', sha256(String(data.newPassword)));
      return jsonResponse({ success: true });
    }

    if (data.action === 'update_phone') {
      const s = getSheet('RDV');
      s.getRange(parseInt(data.row), 4).setValue(data.telephone || '');
      return jsonResponse({ success: true });
    }

    if (data.action === 'mark_termine') {
      const s = getSheet('RDV');
      s.getRange(parseInt(data.row), 11).setValue(data.termine ? true : false);
      return jsonResponse({ success: true });
    }

    if (data.action === 'update') {
      const s = getSheet('RDV');
      const row = parseInt(data.row);
      s.getRange(row, 8).setValue(data.status);
      if (data.status === 'Annulé par cliente') {
        const rows = s.getDataRange().getValues();
        const r = rows[row - 1];
        notifyTelegramAnnulation({ prenom: r[1], nom: r[2], telephone: r[3], prestation: r[4], date: r[5], creneau: r[6] });
      }
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
    const blacklistVal = config.getRange('C1').getValue();
    const clientNotesVal = config.getRange('D1').getValue();

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
        message: rows[i][8] || '',
        reference: rows[i][9] || '',
        termine: rows[i][10] === true
      });
    }

    return jsonResponse({
      success: true,
      bookings: bookings,
      disponibilites: dispoVal ? JSON.parse(dispoVal) : {},
      tarifs: tarifsVal ? JSON.parse(tarifsVal) : null,
      blacklist: blacklistVal ? JSON.parse(blacklistVal) : [],
      clientNotes: clientNotesVal ? JSON.parse(clientNotesVal) : {}
    });

  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function notifyTelegramNewRdv(data) {
  const lines = [
    '📅 Nouvelle demande de RDV',
    `${data.prenom || ''} ${data.nom || ''}`.trim(),
    data.prestation ? `Prestation : ${data.prestation}` : '',
    data.date ? `Date : ${data.date}` : '',
    data.creneau ? `Créneau : ${data.creneau}` : '',
    data.telephone ? `Téléphone : ${data.telephone}` : '',
    data.message ? `Message : ${data.message}` : '',
    data.reference ? `Référence : ${data.reference}` : ''
  ].filter(Boolean);
  sendTelegramMessage(lines.join('\n'));
}

function notifyTelegramAnnulation(data) {
  const lines = [
    '❌ Annulation RDV par cliente',
    `${data.prenom || ''} ${data.nom || ''}`.trim(),
    data.prestation ? `Prestation : ${data.prestation}` : '',
    data.date ? `Date : ${data.date}` : '',
    data.creneau ? `Créneau : ${data.creneau}` : '',
    data.telephone ? `Téléphone : ${data.telephone}` : ''
  ].filter(Boolean);
  sendTelegramMessage(lines.join('\n'));
}

function sha256(txt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(txt), Utilities.Charset.UTF_8);
  return bytes.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function normaliserTelServeur(tel) {
  const digits = String(tel == null ? '' : tel).replace(/\D/g, '');
  return digits.slice(-9); // les 9 derniers chiffres (ignore le 0 initial ou +33)
}

function estDansListeNoire(tel) {
  try {
    const val = getSheet('Config').getRange('C1').getValue();
    const liste = val ? JSON.parse(val) : [];
    const t = normaliserTelServeur(tel);
    if (!t) return false;
    return liste.some(function(n) { return normaliserTelServeur(n) === t; });
  } catch (err) {
    return false;
  }
}

function getDateKey(value) {
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, 'Europe/Paris', 'yyyy-MM-dd');
  return String(value).split('T')[0];
}

function sendTelegramMessage(text) {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('TELEGRAM_BOT_TOKEN');
  const chatId = props.getProperty('TELEGRAM_CHAT_ID');
  if (!token || !chatId) { Logger.log('TELEGRAM_BOT_TOKEN ou TELEGRAM_CHAT_ID manquant dans les Script Properties'); return; }
  try {
    const res = UrlFetchApp.fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: text }),
      muteHttpExceptions: true
    });
    Logger.log('Réponse Telegram (code ' + res.getResponseCode() + ') : ' + res.getContentText());
  } catch (err) {
    Logger.log('Erreur notification Telegram: ' + err);
  }
}

// Fonction de test : sélectionne "testerTelegram" dans le menu déroulant en haut de l'éditeur puis clique ▶ Exécuter.
// Regarde ensuite "Exécutions" (icône horloge à gauche) pour voir le résultat détaillé.
function testerTelegram() {
  sendTelegramMessage('✅ Test de notification — si tu reçois ce message, tout fonctionne !');
}

function sendDailyReminders() {
  const tz = 'Europe/Paris';
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowKey = Utilities.formatDate(tomorrow, tz, 'yyyy-MM-dd');

  const rdvSheet = getSheet('RDV');
  const rows = rdvSheet.getDataRange().getValues();
  const rdvsDemain = [];
  for (let i = 1; i < rows.length; i++) {
    const status = rows[i][7] || 'En attente';
    if (status !== 'Accepté') continue;
    if (getDateKey(rows[i][5]) !== tomorrowKey) continue;
    rdvsDemain.push({
      prenom: rows[i][1],
      nom: rows[i][2],
      telephone: rows[i][3],
      prestation: rows[i][4],
      creneau: rows[i][6]
    });
  }

  if (!rdvsDemain.length) return;

  const lines = ['⏰ Rappel : RDV demain (' + tomorrowKey + ')', ''];
  rdvsDemain.forEach(b => {
    lines.push(`• ${b.creneau || '?'} — ${b.prenom} ${b.nom} (${b.prestation})${b.telephone ? ' — 📞 ' + b.telephone : ''}`);
  });

  sendTelegramMessage(lines.join('\n'));
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
