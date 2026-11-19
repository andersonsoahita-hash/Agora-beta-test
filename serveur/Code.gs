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

/* ================= NOTIFICATIONS (OneSignal) =================
   Réglages > Propriétés du script : OS_APP_ID (App ID) et OS_API_KEY (clé « REST API Key »).
   Lancer UNE fois « installerNotifs » : crée 3 envois automatiques par jour.
   L'appli envoie à OneSignal deux étiquettes : last_day (dernier jour d'entraînement, UTC) et streak. */
var APP_URL_ = 'https://andersonsoahita-hash.github.io/Agora-beta-test/';
function jourUTC_(decalage) { return Utilities.formatDate(new Date(Date.now() + (decalage || 0) * 86400000), 'UTC', 'yyyy-MM-dd'); }
function notif_(titre, texte, filtres) {
  var P = PropertiesService.getScriptProperties(), id = P.getProperty('OS_APP_ID'), cle = P.getProperty('OS_API_KEY');
  if (!id || !cle) { Logger.log('OS_APP_ID / OS_API_KEY manquants'); return; }
  var r = UrlFetchApp.fetch('https://api.onesignal.com/notifications', { method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Key ' + cle },
    payload: JSON.stringify({ app_id: id, target_channel: 'push', url: APP_URL_, filters: filtres, headings: { fr: titre, en: titre }, contents: { fr: texte, en: texte }, ttl: 43200 }) });
  Logger.log(r.getResponseCode() + ' ' + r.getContentText().slice(0, 200));
}
var DEFIS_ = [{"n":"Argumenter une idée","s":["Tout le monde devrait apprendre à parler en public","Le téléphone devrait être interdit en réunion","Le luxe est un art avant d’être un commerce","Madagascar a un atout sous-estimé","Travailler à distance rend meilleur","Il faut échouer pour apprendre","Un bon leader sait écouter avant de parler","Les études ne suffisent plus"]},{"n":"Parler à partir d’un mot","s":["Confiance","Héritage","Rareté","Patience","Frontière","Lenteur","Élégance","Risque","Mémoire","Silence"]},{"n":"Raconter une histoire","s":["Une fois où j’ai eu peur de parler","Un échec qui m’a appris quelque chose","La rencontre qui a changé mon parcours","Un voyage qui m’a marqué","Mon premier jour dans un lieu inconnu"]},{"n":"Défendre un camp","s":["Les émissions de télé-réalité devraient être diffusées après 22 h","L’intelligence artificielle devrait être autorisée dans la création artistique","La semaine de travail devrait durer quatre jours","Le luxe peut être durable","La prise de parole en public devrait être obligatoire à l’école","Les réseaux sociaux devraient être interdits avant 15 ans","Le télétravail devrait être un droit","Les marques de luxe devraient publier leurs prix de fabrication","Il faut rendre le vote obligatoire","L’uniforme devrait être obligatoire à l’école","La publicité devrait être interdite dans l’espace public","Les stages non rémunérés devraient être interdits","Il faut limiter le tourisme de masse","Les études supérieures devraient être gratuites","Les influenceurs devraient être soumis aux mêmes règles que les médias","Les cryptomonnaies sont l’avenir de l’argent","Il faut apprendre à coder dès le primaire","La seconde main devrait remplacer le neuf dans la mode","Les entreprises devraient imposer des quotas de diversité","Le salaire de chacun devrait être public","Le téléphone devrait être interdit en réunion","Les notes à l’école devraient être supprimées","Il faut taxer davantage les produits de luxe","La contrefaçon fait du bien aux grandes marques","Madagascar devrait miser d’abord sur le tourisme","Les jeunes diplômés devraient d’abord travailler à l’étranger","Le service civique devrait être obligatoire","Les voitures devraient être interdites en centre-ville","On devrait pouvoir noter son employeur publiquement","L’expérience compte plus que le diplôme","Les musées devraient être gratuits","Le bénévolat devrait compter dans les études","Il faut limiter les écrans à la maison","Le client n’a pas toujours raison","La concurrence est plus utile que la coopération"]},{"n":"Répondre à une question","s":["Pourquoi devrions-nous vous écouter ?","Quelle est votre plus grande faiblesse ?","Qu’est-ce qui vous différencie ?","Pourquoi ce projet maintenant ?","Et si vous vous trompiez ?"]},{"n":"Relier ses idées (5 étapes)","s":["Qu’est-ce que le bonheur ?","Faut-il toujours dire la vérité ?","La réussite se mérite-t-elle ?","Parler en public est-il un don ?","Faut-il apprendre à perdre ?","Le luxe est-il utile ?","Que vaut une promesse ?"]},{"n":"Discours complet","s":["Pourquoi j’ai écrit ce livre","Ce que je veux transmettre à ma génération","Présenter un projet à un jury","Un discours de remerciement"]},{"n":"Recherche express","s":[]},{"n":"Discours à jouer","s":[]}]; /* mêmes exercices que l'appli, dans le même ordre */
function pick_(a, n) { return a[((n % a.length) + a.length) % a.length]; }
function numJour_() { return Math.floor((Date.now() + 3600000) / 86400000); } /* même formule que l'appli (heure de Paris) */
function notifDefi() { // 18 h : le vrai défi du jour, seulement à ceux qui ne se sont pas encore entraînés aujourd'hui
  var n = numJour_(), t = pick_(DEFIS_, n), suj = t.s.length ? t.s[n % t.s.length] : '';
  var txt = suj ? [t.n + ' : « ' + suj + ' ». Quelques minutes suffisent.', 'Au programme ce soir : « ' + suj + ' » Lance-toi.', 'Ton défi : ' + t.n.toLowerCase() + ', sujet « ' + suj + ' ».'] : [t.n + ' : ton défi du jour t’attend.', 'Ce soir, ' + t.n.toLowerCase() + '. Quelques minutes suffisent.'];
  notif_(pick_(['Ton défi du jour', 'Défi du jour', 'Prêt pour ton défi ?'], n), pick_(txt, n), [{ field: 'tag', key: 'ab', relation: '=', value: '1' }, { field: 'tag', key: 'last_day', relation: '!=', value: jourUTC_(0) }]);
}
function notifSerie() { // 20 h : série en jeu (entraîné hier, pas encore aujourd'hui, série d'au moins 2 jours)
  var n = numJour_(); notif_(pick_(['Ta série est en jeu', 'Ne casse pas ta série', 'Encore une séance ?'], n), pick_(['Une courte séance ce soir et ta série continue.', 'Deux minutes de parole suffisent pour garder ta série.', 'Tu t’es entraîné hier : continue ce soir, ça compte.'], n), [{ field: 'tag', key: 'last_day', relation: '=', value: jourUTC_(-1) }, { field: 'tag', key: 'streak', relation: '>', value: '1' }]);
}
function notifRelance() { // 12 h : relance après 3 jours sans s'entraîner
  var n = numJour_(); notif_(pick_(['Ta voix te manque', 'On t’attend sur Agora', 'Reprends la parole'], n), pick_(['Ça fait trois jours. Reprends avec un exercice court.', 'Un exercice de deux minutes suffit pour reprendre le rythme.', 'Ton dernier entraînement date de trois jours : une petite séance ?'], n), [{ field: 'tag', key: 'last_day', relation: '=', value: jourUTC_(-3) }]);
}

function notifPaliers() { // 21 h : un palier de série atteint aujourd'hui (3, 7, 14, 30, 60, 100 jours)
  [3, 7, 14, 30, 60, 100].forEach(function (n) {
    notif_(n + ' jours de suite !', pick_(['Série de ' + n + ' jours : bravo, ta régularité paie.', n + ' jours d’affilée. Continue, tu progresses vraiment.'], n),
      [{ field: 'tag', key: 'streak', relation: '=', value: String(n) }, { field: 'tag', key: 'last_day', relation: '=', value: jourUTC_(0) }]);
  });
}
function notifVersion() { // toutes les heures : prévient quand une nouvelle version de l'appli est en ligne (1 fois par jour au maximum)
  var P = PropertiesService.getScriptProperties(), r = UrlFetchApp.fetch(APP_URL_ + 'sw.js?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return;
  var m = /abeta-([0-9a-f]+)/.exec(r.getContentText()); if (!m) return;
  var vue = P.getProperty('VERSION_VUE'), der = P.getProperty('VERSION_ENVOI') || '';
  if (!vue) { P.setProperty('VERSION_VUE', m[1]); return; }
  if (vue === m[1]) return;
  P.setProperty('VERSION_VUE', m[1]);
  if (der === jourUTC_(0)) return;
  P.setProperty('VERSION_ENVOI', jourUTC_(0));
  notif_('Mise à jour disponible', 'Une nouvelle version d’Agora test est prête. Ouvre l’appli pour en profiter.', [{ field: 'tag', key: 'ab', relation: '=', value: '1' }]);
}
function installerNotifs() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (['notifDefi', 'notifSerie', 'notifRelance', 'notifPaliers', 'notifVersion'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('notifRelance').timeBased().everyDays(1).atHour(12).inTimezone('Europe/Paris').create();
  ScriptApp.newTrigger('notifDefi').timeBased().everyDays(1).atHour(18).inTimezone('Europe/Paris').create();
  ScriptApp.newTrigger('notifSerie').timeBased().everyDays(1).atHour(20).inTimezone('Europe/Paris').create();
  ScriptApp.newTrigger('notifPaliers').timeBased().everyDays(1).atHour(21).inTimezone('Europe/Paris').create();
  ScriptApp.newTrigger('notifVersion').timeBased().everyHours(1).create();
}
