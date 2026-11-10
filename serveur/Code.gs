/* Agora test · serveur (Google Apps Script) — by Stephen
   À coller dans Extensions > Apps Script du Google Sheet. */

var FEUILLES = {
  comptes: ['Date', 'Nom', 'Prénoms', 'Email', 'Téléphone', 'Âge', 'Statut', 'Code d’activation', 'Expire le', 'Dernière connexion', 'Séances', 'Appareil', 'Mot de passe (empreinte)', 'Session'],
  avis: ['Date', 'Email', 'Écran', 'Avis'],
  journal: ['Date', 'Action'],
  usage: ['Date', 'Email', 'Fonction']
};
var NCOL = 14;
var APP_URL = 'https://andersonsoahita-hash.github.io/Agora-beta-test/'; // adresse de l'appli de test

function feuille_(nom) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(nom);
  if (!sh) { sh = ss.insertSheet(nom); sh.appendRow(FEUILLES[nom.toLowerCase()]); sh.setFrozenRows(1); }
  if (nom === 'Comptes') sh.getRange(1, 1, 1, NCOL).setValues([FEUILLES.comptes]); // garde les titres à jour
  return sh;
}
function installer() { // à lancer une fois : crée les onglets
  ['Comptes', 'Avis', 'Journal', 'Usage'].forEach(feuille_);
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Feuille 1'); if (s && s.getLastRow() === 0) s.setName('Notes');
  if (!PropertiesService.getScriptProperties().getProperty('CLE_ADMIN')) PropertiesService.getScriptProperties().setProperty('CLE_ADMIN', Utilities.getUuid());
}
function journal_(a) { feuille_('Journal').appendRow([new Date(), a]); }
function alea_(n, alphabet) { var s = ''; for (var i = 0; i < n; i++) s += alphabet.charAt(Math.floor(Math.random() * alphabet.length)); return s; }
function code_() { return alea_(8, 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'); } // code d'activation lisible
function jeton_() { return Utilities.getUuid().replace(/-/g, ''); }
function hash_(mdp, email) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, email + '|' + mdp + '|agora-test')
    .map(function (x) { return ('0' + (x & 255).toString(16)).slice(-2); }).join('');
}

/* ---------- menu dans le Sheet ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Agora test')
    .addItem('Valider / renvoyer un code (ligne sélectionnée)', 'valider')
    .addItem('Bloquer la ligne sélectionnée', 'bloquer')
    .addItem('Débloquer la ligne sélectionnée', 'debloquer')
    .addItem('Supprimer le compte de la ligne sélectionnée', 'supprimer')
    .addSeparator().addItem('Afficher ma clé admin', 'montrerCle').addToUi();
}
function emailLigne_() {
  var sh = SpreadsheetApp.getActiveSheet();
  if (sh.getName() !== 'Comptes') throw new Error('Va dans l’onglet Comptes et sélectionne une ligne.');
  var r = sh.getActiveRange().getRow(); if (r < 2) throw new Error('Sélectionne une ligne de compte.');
  return String(sh.getRange(r, 4).getValue());
}
function valider() { var r = agir_('valider', emailLigne_()); SpreadsheetApp.getActive().toast(r.ok ? 'Code envoyé par e-mail' : r.erreur); }
function bloquer() { agir_('bloquer', emailLigne_()); }
function debloquer() { agir_('debloquer', emailLigne_()); }
function supprimer() { agir_('supprimer', emailLigne_()); }
function montrerCle() { SpreadsheetApp.getUi().alert('Clé admin (pour la page de suivi) :\n' + PropertiesService.getScriptProperties().getProperty('CLE_ADMIN')); }

/* ---------- outils ---------- */
function rep_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function trouver_(champ, valeur) { // renvoie {sh, r, v} de la ligne dont la colonne (index) vaut valeur
  var sh = feuille_('Comptes'), n = sh.getLastRow(); if (n < 2 || !valeur) return null;
  var d = sh.getRange(2, 1, n - 1, NCOL).getValues();
  for (var i = 0; i < d.length; i++) if (String(d[i][champ]).toLowerCase() === String(valeur).toLowerCase()) return { sh: sh, r: i + 2, v: d[i] };
  return null;
}
function etat_(v) {
  var st = String(v[6]);
  if (st === 'validé' && v[8] && new Date(v[8]) < new Date(new Date().toDateString())) return 'expiré';
  return st;
}

/* ---------- actions admin ---------- */
function agir_(op, email, date) {
  var t = trouver_(3, email); if (!t) return { ok: false, erreur: 'Compte introuvable' };
  var sh = t.sh, r = t.r, v = t.v;
  if (op === 'valider') { // (re)donne un code d'activation à usage unique et remet le compte à zéro
    var code = code_();
    sh.getRange(r, 7).setValue('validé'); sh.getRange(r, 8).setValue(code);
    sh.getRange(r, 12).setValue(''); sh.getRange(r, 13).setValue(''); sh.getRange(r, 14).setValue('');
    MailApp.sendEmail({ to: v[3], subject: 'Ton code d’accès Agora (test)',
      body: 'Bonjour ' + v[2] + ',\n\nTa demande est validée.\n\n1. Ouvre Agora sur l’appareil que tu vas utiliser (un seul appareil par compte) :\n' + APP_URL + '\n   Sur téléphone, tu peux l’ajouter à l’écran d’accueil.\n2. Choisis « J’ai un code », puis entre ton e-mail et ce code d’activation :\n\n      ' + code + '\n\n3. Crée ton mot de passe. Il te servira à te reconnecter.\n\nCe code ne marche qu’une seule fois. Ne le partage pas.\n\nStephen' });
    journal_('Code d’activation envoyé à ' + v[3]);
  } else if (op === 'bloquer') { sh.getRange(r, 7).setValue('bloqué'); journal_('Bloqué : ' + v[3]); }
  else if (op === 'debloquer') { sh.getRange(r, 7).setValue('validé'); journal_('Débloqué : ' + v[3]); }
  else if (op === 'supprimer') { sh.getRange(r, 7).setValue('supprimé'); [8, 12, 13, 14].forEach(function (c) { sh.getRange(r, c).setValue(''); }); journal_('Supprimé : ' + v[3]); }
  else if (op === 'appareil') { sh.getRange(r, 12).setValue(''); sh.getRange(r, 14).setValue(''); journal_('Appareil réinitialisé : ' + v[3]); }
  else if (op === 'expire') { sh.getRange(r, 9).setValue(date || ''); journal_('Fin d’accès ' + (date || 'retirée') + ' : ' + v[3]); }
  else if (op === 'effacer') { // suppression définitive : la ligne et son suivi disparaissent (ses avis sont gardés)
    sh.deleteRow(r);
    var su = feuille_('Usage'), nu = su.getLastRow();
    if (nu > 1) { var du = su.getRange(2, 1, nu - 1, 3).getValues(); for (var i = du.length - 1; i >= 0; i--) if (String(du[i][1]).toLowerCase() === String(v[3]).toLowerCase()) su.deleteRow(i + 2); }
    journal_('Compte effacé définitivement : ' + v[3]);
  }
  else return { ok: false, erreur: 'Action inconnue' };
  return { ok: true };
}

/* ---------- API ---------- */
function admin_(b) { return b.cle === PropertiesService.getScriptProperties().getProperty('CLE_ADMIN'); }
function doPost(e) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var b = JSON.parse(e.postData.contents || '{}'), a = b.action;
    var em = String(b.email || '').trim().toLowerCase();

    if (a === 'demande') {
      if (!/^\S+@\S+\.\S+$/.test(em) || !b.nom || !b.prenoms) return rep_({ ok: false, erreur: 'Champs manquants' });
      var ex = trouver_(3, em);
      if (ex && etat_(ex.v) !== 'supprimé') return rep_({ ok: true, etat: etat_(ex.v) });
      if (ex) { ex.sh.deleteRow(ex.r); }
      feuille_('Comptes').appendRow([new Date(), b.nom, b.prenoms, em, b.tel || '', b.age || '', 'en attente', '', '', '', 0, '', '', '']);
      journal_('Nouvelle demande : ' + em);
      return rep_({ ok: true, etat: 'en attente' });
    }

    if (a === 'activer') { // code à usage unique + création du mot de passe
      var t = trouver_(3, em);
      if (!t || etat_(t.v) !== 'validé') return rep_({ ok: false, erreur: 'Compte non validé.' });
      if (!t.v[7] || String(t.v[7]).toUpperCase() !== String(b.code || '').trim().toUpperCase()) return rep_({ ok: false, erreur: 'Code invalide ou déjà utilisé.' });
      if (String(b.mdp || '').length < 6) return rep_({ ok: false, erreur: 'Mot de passe : 6 caractères minimum.' });
      var jt = jeton_();
      t.sh.getRange(t.r, 8).setValue(''); // le code est consommé
      t.sh.getRange(t.r, 12).setValue(b.appareil || ''); t.sh.getRange(t.r, 13).setValue(hash_(b.mdp, em)); t.sh.getRange(t.r, 14).setValue(jt);
      t.sh.getRange(t.r, 10).setValue(new Date());
      journal_('Compte activé : ' + em);
      return rep_({ ok: true, etat: 'validé', jeton: jt, prenoms: t.v[2] });
    }

    if (a === 'connexion') { // reconnexion avec le mot de passe, même appareil
      var c = trouver_(3, em);
      if (!c || !c.v[12] || c.v[12] !== hash_(String(b.mdp || ''), em)) return rep_({ ok: false, erreur: 'E-mail ou mot de passe incorrect.' });
      var s = etat_(c.v); if (s !== 'validé') return rep_({ ok: true, etat: s });
      if (c.v[11] && c.v[11] !== b.appareil) return rep_({ ok: true, etat: 'autre appareil' });
      var j2 = jeton_();
      c.sh.getRange(c.r, 12).setValue(b.appareil || ''); c.sh.getRange(c.r, 14).setValue(j2); c.sh.getRange(c.r, 10).setValue(new Date());
      return rep_({ ok: true, etat: 'validé', jeton: j2, prenoms: c.v[2] });
    }

    if (a === 'etat') { // vérification à chaque ouverture
      var u = trouver_(3, em);
      if (!u) return rep_({ ok: true, etat: 'inconnu' });
      var et = etat_(u.v);
      if (et === 'validé') {
        if (!u.v[13] || u.v[13] !== b.jeton || (u.v[11] && u.v[11] !== b.appareil)) return rep_({ ok: true, etat: 'session' });
        u.sh.getRange(u.r, 10).setValue(new Date());
      }
      return rep_({ ok: true, etat: et, prenoms: u.v[2] });
    }

    if (a === 'avis') {
      var w = trouver_(3, em); if (!w || etat_(w.v) !== 'validé' || w.v[13] !== b.jeton) return rep_({ ok: false });
      feuille_('Avis').appendRow([new Date(), em, b.ecran || '', String(b.texte || '').slice(0, 2000)]);
      return rep_({ ok: true });
    }

    if (a === 'usage') { // le testeur envoie les fonctions qu'il a utilisées
      var q = trouver_(3, em); if (!q || etat_(q.v) !== 'validé' || q.v[13] !== b.jeton) return rep_({ ok: false });
      var evs = (b.events || []).slice(0, 300).map(function (x) { return [new Date(+x.t || Date.now()), em, String(x.e || '').slice(0, 80)]; });
      if (evs.length) { var su2 = feuille_('Usage'); su2.getRange(su2.getLastRow() + 1, 1, evs.length, 3).setValues(evs); }
      return rep_({ ok: true });
    }

    if (a === 'liste') {
      if (!admin_(b)) return rep_({ ok: false, erreur: 'Clé invalide' });
      var sh = feuille_('Comptes'), n = sh.getLastRow(), d = n > 1 ? sh.getRange(2, 1, n - 1, NCOL).getValues() : [];
      var sa = feuille_('Avis'), na = sa.getLastRow(), da = na > 1 ? sa.getRange(2, 1, na - 1, 4).getValues() : [];
      var sj = feuille_('Journal'), nj = sj.getLastRow(), dj = nj > 1 ? sj.getRange(Math.max(2, nj - 39), 1, Math.min(40, nj - 1), 2).getValues() : [];
      var su3 = feuille_('Usage'), nu3 = su3.getLastRow(), du3 = nu3 > 1 ? su3.getRange(Math.max(2, nu3 - 4999), 1, Math.min(5000, nu3 - 1), 3).getValues() : [], usage = {};
      du3.forEach(function (x) { var m = String(x[1]).toLowerCase(), f = String(x[2]); usage[m] = usage[m] || {}; usage[m][f] = (usage[m][f] || 0) + 1; });
      return rep_({ ok: true, usage: usage,
        comptes: d.map(function (x) { return { nom: x[1] + ' ' + x[2], email: x[3], tel: x[4], age: x[5], demande: x[0], statut: etat_(x), derniere: x[9], expire: x[8], appareil: !!x[11], active: !!x[12], code: !!x[7] }; }),
        avis: da, journal: dj });
    }

    if (a === 'admin') {
      if (!admin_(b)) return rep_({ ok: false, erreur: 'Clé invalide' });
      return rep_(agir_(b.op, em, b.date));
    }
    return rep_({ ok: false, erreur: 'Action inconnue' });
  } catch (err) { return rep_({ ok: false, erreur: String(err) }); }
  finally { lock.releaseLock(); }
}
function doGet() { return rep_({ ok: true, service: 'Agora test' }); }
