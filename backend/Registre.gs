/**
 * REGISTRE D'ACCIDENTS, D'INCIDENTS ET DE PREMIERS SECOURS — partie serveur
 * Projet Apps Script « Inspection API » (Sheet « Inspection usine »).
 *
 * Copie de référence : ce bloc est ajouté À LA FIN de Code.gs dans l'éditeur Apps Script.
 * Il utilise getSheet() de Code.gs. Aiguillage ajouté dans Code.gs :
 *   - doPost (ligne « try { var __ab = JSON.parse(...) ») :
 *       if (__ab && __ab.action === "registre") { return __handleRegistre(__ab); }
 *       if (__ab && __ab.action === "registreMesure") { return __handleRegistreMesure(__ab); }
 *   - doGet, juste après « const action = ... » :
 *       if (action === 'getRegistre') { return __getRegistre(); }
 * Puis redéployer une NOUVELLE VERSION du déploiement AKfycbxlk41R7pJD6WBJ… (voir le guide technique).
 */

var REG_SHEET = 'Registre';
var REG_DEST = 'jtremblay@ntcamps.com, jpronovost@ntcamps.com';
var REG_HEADERS = ['ID', 'Horodatage', 'Date', 'Heure', 'Établissement', 'Lieu', 'Métier ou fonction',
  'Description et causes', 'Blessure ou malaise', 'Premiers secours et transport',
  'Personne secourue', 'Absente au-delà de la journée', 'Signature personne secourue',
  'Secouriste', 'Signature secouriste', 'Mesures correctives', 'Personne responsable',
  'Date de mise en place', 'Inscrit par', 'Mesure complétée par', 'Mesure complétée le'];
var REG_MAX_CELL = 49000; // limite Google Sheets : 50 000 caractères par cellule

function __regSheet() {
  var sh = getSheet(REG_SHEET); // fonction de Code.gs (crée l'onglet au besoin)
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, REG_HEADERS.length).setValues([REG_HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    // Dates et heures gardées en texte pour éviter les décalages de fuseau horaire
    sh.getRange('C:D').setNumberFormat('@');
    sh.getRange('R:R').setNumberFormat('@');
  }
  return sh;
}

function __regJson(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function __regCell(v) {
  v = (v == null) ? '' : String(v);
  return v.length > REG_MAX_CELL ? '(trop volumineuse pour une cellule)' : v;
}

function __handleRegistre(data) {
  var e = data.entree || {};
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = __regSheet();
    sh.appendRow([
      e.id, new Date(), e.date, e.heure, e.etablissement, e.lieu, e.metier,
      e.description, e.blessure, e.secours, e.secouruNom, e.absent === 'oui' ? 'Oui' : 'Non',
      __regCell(e.sigSecouru), e.secouristeNom, __regCell(e.sigSecouriste),
      e.mesures, e.responsable, e.dateMesure, e.inscritPar, '', ''
    ]);
  } catch (err) {
    console.error('Registre : écriture impossible', err);
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
  // Courriel avec le PDF de l'entrée
  try {
    var attach = [];
    if (data.pdfBase64) {
      attach.push(Utilities.newBlob(Utilities.base64Decode(data.pdfBase64), 'application/pdf', data.filename || ('Registre_' + e.id + '.pdf')));
    }
    var html = '<p>Nouvelle entrée au <b>registre d’accidents, d’incidents et de premiers secours</b> de l’usine.</p>'
      + '<table cellpadding="4" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px">'
      + __regRow('Entrée', e.id) + __regRow('Date et heure', (e.date || '') + (e.heure ? ' à ' + e.heure : ''))
      + __regRow('Lieu', e.lieu) + __regRow('Personne secourue', e.secouruNom + (e.metier ? ' (' + e.metier + ')' : ''))
      + __regRow('Blessure ou malaise', e.blessure) + __regRow('Premiers secours et transport', e.secours)
      + __regRow('Secouriste', e.secouristeNom) + __regRow('Mesures correctives', e.mesures || 'À compléter')
      + __regRow('Inscrit par', e.inscritPar) + '</table>'
      + '<p style="color:#777;font-size:12px">Le PDF de l’entrée est en pièce jointe.</p>';
    MailApp.sendEmail({ to: REG_DEST, name: 'Registre d’accidents – Usine', subject: data.subject || ('Registre d’accidents – ' + e.id), htmlBody: html, attachments: attach });
  } catch (err) {
    console.error('Registre : courriel non envoyé', err);
  }
  return __regJson({ ok: true, id: e.id });
}

function __regRow(label, val) {
  var esc = String(val == null ? '' : val).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\n/g, '<br>');
  return '<tr><td style="color:#777;vertical-align:top;white-space:nowrap">' + label + '</td><td>' + (esc || '—') + '</td></tr>';
}

function __handleRegistreMesure(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = __regSheet();
    var ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (var i = 1; i < ids.length; i++) {
      if (String(ids[i][0]) === String(data.id)) {
        var r = i + 1;
        sh.getRange(r, 16, 1, 3).setValues([[data.mesures || '', data.responsable || '', data.dateMesure || '']]);
        sh.getRange(r, 20, 1, 2).setValues([[data.par || '', new Date()]]);
        return __regJson({ ok: true });
      }
    }
    console.error('Registre : entrée introuvable pour la mesure', data.id);
    return __regJson({ ok: false, error: 'introuvable' });
  } catch (err) {
    console.error('Registre : mise à jour impossible', err);
    return __regJson({ ok: false });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function __getRegistre() {
  var sh = __regSheet();
  var last = sh.getLastRow();
  if (last < 2) return __regJson({ registre: [] });
  var tz = Session.getScriptTimeZone();
  function d(v) { return (v instanceof Date) ? Utilities.formatDate(v, tz, 'yyyy-MM-dd') : String(v || ''); }
  function h(v) { return (v instanceof Date) ? Utilities.formatDate(v, tz, 'HH:mm') : String(v || ''); }
  var rows = sh.getRange(2, 1, last - 1, REG_HEADERS.length).getValues();
  var list = rows.filter(function (x) { return x[0]; }).map(function (x) {
    // Les signatures ne sont pas renvoyées (trop lourdes pour la liste)
    return {
      id: String(x[0]), date: d(x[2]), heure: h(x[3]), etablissement: x[4], lieu: x[5], metier: x[6],
      description: x[7], blessure: x[8], secours: x[9], secouruNom: x[10], absent: x[11] === 'Oui' ? 'oui' : 'non',
      secouristeNom: x[13], mesures: x[15], responsable: x[16], dateMesure: d(x[17]), inscritPar: x[18], mesurePar: x[19]
    };
  });
  return __regJson({ registre: list });
}
