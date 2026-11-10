/* Agora test · serveur (Google Apps Script) — by Stephen
   À coller dans Extensions > Apps Script du Google Sheet. Voir LISEZMOI.md. */

var FEUILLES = {
  comptes: ['Date', 'Nom', 'Prénoms', 'Email', 'Téléphone', 'Âge', 'Statut', 'Code', 'Expire le', 'Dernière connexion', 'Séances', 'Appareil'],
  avis: ['Date', 'Email', 'Écran', 'Avis'],
  journal: ['Date', 'Action']
};
var APP_URL = 'https://andersonsoahita-hash.github.io/Agora-beta-test/'; // adresse de l'appli de test

function feuille_(nom) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(nom);
  if (!sh) { sh = ss.insertSheet(nom); sh.appendRow(FEUILLES[nom.toLowerCase()] || []); sh.setFrozenRows(1); }
  return sh;
}
function installer() { // à lancer une fois : crée les onglets
  ['Comptes', 'Avis', 'Journal'].forEach(feuille_);
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Feuille 1'); if (s && s.getLastRow() === 0) s.setName('Notes');
  if (!PropertiesService.getScriptProperties().getProperty('CLE_ADMIN')) PropertiesService.getScriptProperties().setProperty('CLE_ADMIN', Utilities.getUuid());
}
function journal_(a) { feuille_('Journal').appendRow([new Date(), a]); }
function code_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); }

/* ---------- menu dans le Sheet ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Agora test')
    .addItem('Valider la ligne sélectionnée (envoie l’accès)', 'valider')
    .addItem('Bloquer la ligne sélectionnée', 'bloquer')
    .addItem('Débloquer la ligne sélectionnée', 'debloquer')
    .addItem('Supprimer le compte de la ligne sélectionnée', 'supprimer')
    .addSeparator().addItem('Afficher ma clé admin', 'montrerCle').addToUi();
}
function ligne_() {
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== 'Comptes') throw new Error('Va dans l’onglet Comptes et sélectionne une ligne.');
  var r = sh.getActiveRange().getRow(); if (r < 2) throw new Error('Sélectionne une ligne de compte.');
  return { sh: sh, r: r, v: sh.getRange(r, 1, 1, 12).getValues()[0] };
}
function valider() {
  var L = ligne_(), v = L.v, code = v[7] || code_();
  L.sh.getRange(L.r, 7).setValue('validé'); L.sh.getRange(L.r, 8).setValue(code);
  var lien = APP_URL + '?c=' + code;
  MailApp.sendEmail({ to: v[3], subject: 'Ton accès à Agora (test)',
    body: 'Bonjour ' + v[2] + ',\n\nTa demande est validée. Ouvre ce lien sur ton téléphone ou ton ordinateur (un seul appareil par compte) :\n' + lien + '\n\nCe lien est personnel : ne le partage pas.\n\nStephen' });
  journal_('Accès validé et envoyé à ' + v[3]);
  SpreadsheetApp.getActive().toast('Accès envoyé à ' + v[3]);
}
function bloquer() { var L = ligne_(); L.sh.getRange(L.r, 7).setValue('bloqué'); journal_('Bloqué : ' + L.v[3]); }
function debloquer() { var L = ligne_(); L.sh.getRange(L.r, 7).setValue('validé'); journal_('Débloqué : ' + L.v[3]); }
function supprimer() { var L = ligne_(); L.sh.getRange(L.r, 7).setValue('supprimé'); L.sh.getRange(L.r, 8).setValue(''); journal_('Supprimé : ' + L.v[3]); }
function montrerCle() { SpreadsheetApp.getUi().alert('Clé admin (pour la page de suivi) :\n' + PropertiesService.getScriptProperties().getProperty('CLE_ADMIN')); }

/* ---------- API (appelée par l'appli) ---------- */
function rep_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function trouver_(champ, valeur) { // renvoie {r, v} de la ligne dont la colonne vaut valeur
  var sh = feuille_('Comptes'), n = sh.getLastRow(); if (n < 2) return null;
  var d = sh.getRange(2, 1, n - 1, 12).getValues();
  for (var i = 0; i < d.length; i++) if (String(d[i][champ]).toLowerCase() === String(valeur).toLowerCase() && valeur) return { sh: sh, r: i + 2, v: d[i] };
  return null;
}
function etat_(v) {
  var st = String(v[6]);
  if (st === 'validé' && v[8] && new Date(v[8]) < new Date(new Date().toDateString())) return 'expiré';
  return st;
}

/* ---------- actions admin (depuis la page de suivi, avec la clé) ---------- */
function agir_(op, email, date) {
  var t = trouver_(3, email); if (!t) return { ok: false, erreur: 'Compte introuvable' };
  var sh = t.sh, r = t.r, v = t.v;
  if (op === 'valider') {
    var code = v[7] || code_();
    sh.getRange(r, 7).setValue('validé'); sh.getRange(r, 8).setValue(code);
    MailApp.sendEmail({ to: v[3], subject: 'Ton accès à Agora (test)',
      body: 'Bonjour ' + v[2] + ',\n\nTa demande est validée. Ouvre ce lien sur ton téléphone ou ton ordinateur (un seul appareil par compte) :\n' + APP_URL + '?c=' + code + '\n\nCe lien est personnel : ne le partage pas.\n\nStephen' });
    journal_('Accès validé et envoyé à ' + v[3]);
  } else if (op === 'bloquer') { sh.getRange(r, 7).setValue('bloqué'); journal_('Bloqué : ' + v[3]); }
  else if (op === 'debloquer') { sh.getRange(r, 7).setValue('validé'); journal_('Débloqué : ' + v[3]); }
  else if (op === 'supprimer') { sh.getRange(r, 7).setValue('supprimé'); sh.getRange(r, 8).setValue(''); sh.getRange(r, 12).setValue(''); journal_('Supprimé : ' + v[3]); }
  else if (op === 'appareil') { sh.getRange(r, 12).setValue(''); journal_('Appareil réinitialisé : ' + v[3]); }
  else if (op === 'expire') { sh.getRange(r, 9).setValue(date || ''); journal_('Fin d’accès ' + (date || 'retirée') + ' : ' + v[3]); }
  else return { ok: false, erreur: 'Action inconnue' };
  return { ok: true };
}
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var b = JSON.parse(e.postData.contents || '{}'), a = b.action;
    if (a === 'demande') {
      var em = String(b.email || '').trim().toLowerCase();
      if (!/^\S+@\S+\.\S+$/.test(em) || !b.nom || !b.prenoms) return rep_({ ok: false, erreur: 'Champs manquants' });
      var ex = trouver_(3, em);
      if (ex && etat_(ex.v) !== 'supprimé') return rep_({ ok: true, etat: etat_(ex.v) }); // déjà demandé
      feuille_('Comptes').appendRow([new Date(), b.nom, b.prenoms, em, b.tel || '', b.age || '', 'en attente', '', '', '', 0, '']);
      journal_('Nouvelle demande : ' + em);
      return rep_({ ok: true, etat: 'en attente' });
    }
    if (a === 'etat') { // l'appli demande où en est son accès (code personnel)
      var t = trouver_(7, b.code);
      if (!t) return rep_({ ok: true, etat: 'inconnu' });
      var etat = etat_(t.v);
      if (etat === 'validé') {
        // un seul appareil : le premier qui se présente est enregistré
        if (t.v[11] && b.appareil && t.v[11] !== b.appareil) return rep_({ ok: true, etat: 'autre appareil' });
        if (!t.v[11] && b.appareil) t.sh.getRange(t.r, 12).setValue(b.appareil);
        t.sh.getRange(t.r, 10).setValue(new Date());
        if (b.seances) t.sh.getRange(t.r, 11).setValue(+b.seances || 0);
      }
      return rep_({ ok: true, etat: etat, prenoms: t.v[2], expire: t.v[8] ? Utilities.formatDate(new Date(t.v[8]), 'Europe/Paris', 'yyyy-MM-dd') : '' });
    }
    if (a === 'avis') {
      var c = trouver_(7, b.code); if (!c || etat_(c.v) !== 'validé') return rep_({ ok: false });
      feuille_('Avis').appendRow([new Date(), c.v[3], b.ecran || '', String(b.texte || '').slice(0, 2000)]);
      return rep_({ ok: true });
    }
    if (a === 'liste') { // page de suivi (réservée à Stephen : clé admin)
      if (b.cle !== PropertiesService.getScriptProperties().getProperty('CLE_ADMIN')) return rep_({ ok: false, erreur: 'Clé invalide' });
      var sh = feuille_('Comptes'), n = sh.getLastRow(), d = n > 1 ? sh.getRange(2, 1, n - 1, 12).getValues() : [];
      var sa = feuille_('Avis'), na = sa.getLastRow(), da = na > 1 ? sa.getRange(2, 1, na - 1, 4).getValues() : [];
      return rep_({ ok: true, comptes: d.map(function (x) { return { nom: x[1] + ' ' + x[2], email: x[3], statut: etat_(x), derniere: x[9], seances: x[10], expire: x[8] }; }), avis: da });
    }
    if (a === 'admin') {
      if (b.cle !== PropertiesService.getScriptProperties().getProperty('CLE_ADMIN')) return rep_({ ok: false, erreur: 'Clé invalide' });
      return rep_(agir_(b.op, String(b.email || '').toLowerCase(), b.date));
    }
    return rep_({ ok: false, erreur: 'Action inconnue' });
  } catch (err) { return rep_({ ok: false, erreur: String(err) }); }
  finally { lock.releaseLock(); }
}
function doGet() { return rep_({ ok: true, service: 'Agora test' }); }
