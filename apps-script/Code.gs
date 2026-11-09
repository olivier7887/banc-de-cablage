/**
 * Banc de câblage — comptes élèves et suivi des exercices
 * Script du Google Sheet « Banc de câblage – données ». Créé par Lesenfants Olivier (OL Élec).
 *
 * Ce script a DEUX déploiements (même code) :
 *   1. « Connexion » : Exécuter en tant que : Moi · Qui a accès : utilisateurs de elmarche.be
 *      → doGet : Google indique qui est connecté, le script remet une clé signée à l'app.
 *   2. « Données »   : Exécuter en tant que : Moi · Qui a accès : Tout le monde
 *      → doPost : l'app envoie les résultats, demande un code par e-mail, et le prof gère les accès.
 *
 * Installation : coller ce code dans Extensions › Apps Script, exécuter « installer » une fois,
 * puis créer les deux déploiements (Déployer › Nouveau déploiement › Application Web).
 * Mise à jour : recoller le code, exécuter « installer », puis Gérer les déploiements › crayon ›
 * Nouvelle version, pour CHACUN des deux déploiements (les adresses ne changent pas).
 *
 * Facultatif : dans Paramètres du projet › Propriétés du script, ajouter CODE_PROF = votre code de
 * licence prof. Les professeurs connectés passent alors automatiquement en mode prof.
 */

const DOMAINE = 'elmarche.be';
const APP_URL = 'https://olivier7887.github.io/banc-de-cablage/';
const DUREE_CLE_JOURS = 180;
const CODE_MINUTES = 15;

const ONGLETS = {
  profs: { nom: 'Profs', entetes: ['Adresse', 'Nom', 'Remarque'] },
  classes: { nom: 'Classes', entetes: ['Classe', 'Professeur (adresse)', 'Remarque'] },
  autorises: { nom: 'Autorisés', entetes: ['Adresse ou @domaine', 'Nom', 'Classe', 'Autorisé par', 'Autorisé le', 'Remarque'] },
  demandes: { nom: 'Demandes', entetes: ['Adresse', 'Nom', 'Classe', 'Demandé le', 'Identité vérifiée', 'Statut', 'Traité par', 'Traité le'] },
  resultats: { nom: 'Résultats', entetes: ['Adresse', 'Nom', 'Classe', 'Exercice (clé)', 'Titre', 'Niveau', 'Type', 'Essais', 'Réussi', 'Réussi au essai n°', 'Premier essai', 'Réussi le', 'Dernier essai', 'Meilleure note /20'] },
  journal: { nom: 'Journal', entetes: ['Date', 'Adresse', 'Exercice (clé)', 'Titre', 'Niveau', 'Type', 'Résultat', 'Détail', 'id'] },
};
const CLASSES_INITIALES = ['3P Élec', '3QEM groupe 1', '3QEM groupe 2', '4Q Élec', '4Q Méca'];

/* =========================================================
   Installation (à exécuter une fois depuis l'éditeur)
   ========================================================= */
function installer() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const k in ONGLETS) {
    const o = ONGLETS[k];
    let sh = ss.getSheetByName(o.nom);
    if (!sh) sh = ss.insertSheet(o.nom);
    sh.getRange(1, 1, 1, o.entetes.length).setValues([o.entetes]).setFontWeight('bold').setBackground('#e9edf1');
    sh.setFrozenRows(1);
  }
  const moi = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  const profs = onglet_('profs');
  if (profs.getLastRow() < 2 && moi) profs.appendRow([moi, 'Lesenfants Olivier', 'Propriétaire']);
  const cls = onglet_('classes');
  if (cls.getLastRow() < 2) cls.getRange(2, 1, CLASSES_INITIALES.length, 3).setValues(CLASSES_INITIALES.map(c => [c, moi, '']));
  try {
    cls.getRange('B2:B200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(profs.getRange('A2:A200'), true).setAllowInvalid(false).build());
    onglet_('autorises').getRange('C2:C2000').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(cls.getRange('A2:A200'), true).setAllowInvalid(true).build());
    [['autorises', 1, 300], ['demandes', 1, 300], ['resultats', 1, 280], ['resultats', 5, 300], ['journal', 2, 280], ['journal', 4, 300]].forEach(([k, c, w]) => onglet_(k).setColumnWidth(c, w));
  } catch (e) { /* mise en forme facultative */ }
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SECRET')) props.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  const premier = ss.getSheets()[0];
  if (/^(Feuille 1|Sheet1)$/.test(premier.getName()) && premier.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(premier);
  Logger.log('Installation terminée.');
}

/* =========================================================
   1. Connexion avec le compte de l'école (déploiement réservé à elmarche.be)
   ========================================================= */
function doGet() {
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  let corps;
  try {
    if (!email) throw new Error("Google n'a pas transmis votre adresse. Ouvrez ce lien en étant connecté avec votre compte de l'école.");
    const cle = signer_({ e: email, x: Date.now() + DUREE_CLE_JOURS * 864e5 });
    const st = statut_(email);
    if (st.statut !== 'ok') noterDemande_(email, '', true);
    const n = nomDe_(email);
    corps = `<h1>Bonjour ${html_(n.prenom)} !</h1>
      <p>Vous êtes connecté avec <b>${html_(email)}</b>${st.role === 'prof' ? ' (professeur)' : ''}.</p>
      ${st.statut === 'ok' ? '' : st.statut === 'refuse' ? '<p>Votre accès a été refusé. Parlez-en à votre professeur.</p>' : '<p>Votre demande d’accès est envoyée : votre professeur doit vous autoriser. En attendant, la version d’essai reste disponible.</p>'}
      <a class="btn" href="${APP_URL}#cle=${cle}" target="_top">Ouvrir le Banc de câblage</a>
      <p class="petit">L'app garde votre connexion pendant ${DUREE_CLE_JOURS} jours, même sans internet.</p>`;
  } catch (e) {
    corps = `<h1>Connexion impossible</h1><p>${html_(e.message)}</p>${email ? `<p class="petit">Adresse transmise par Google : <b>${html_(email)}</b></p>` : ''}`;
  }
  return HtmlService.createHtmlOutput(`<!doctype html><meta name="viewport" content="width=device-width,initial-scale=1">
    <style>body{font:16px/1.5 system-ui,sans-serif;background:#eef0ec;color:#1d2321;display:grid;place-items:center;min-height:100vh;margin:0;padding:16px}
    main{background:#fff;border:1px solid #d6dad3;border-radius:14px;padding:28px;max-width:440px}
    h1{margin:0 0 8px;font-size:1.6rem}.btn{display:inline-block;margin:14px 0;background:#d9541e;color:#fff;text-decoration:none;padding:12px 20px;border-radius:8px;font-weight:600}
    .petit{font-size:.85rem;color:#5c6660}</style><main>${corps}</main>`).setTitle('Banc de câblage – connexion');
}

/* =========================================================
   2. Données (déploiement « Tout le monde »), appelé par l'app
   ========================================================= */
function doPost(e) {
  let rep;
  try {
    const q = JSON.parse(e.postData.contents);
    if (PUBLIQUES[q.a]) rep = { ok: true, data: PUBLIQUES[q.a](q) };
    else {
      const email = verifier_(q.cle);
      const fn = ACTIONS[q.a];
      if (!fn) throw new Error('Action inconnue.');
      rep = { ok: true, data: fn(Object.assign({ email }, statut_(email)), q) };
    }
  } catch (err) {
    rep = { ok: false, erreur: err.message, auth: !!err.auth, attente: !!err.attente };
  }
  return ContentService.createTextOutput(JSON.stringify(rep)).setMimeType(ContentService.MimeType.JSON);
}

// sans clé : connexion par code reçu par e-mail
const PUBLIQUES = {
  codeDemander: (q) => codeDemander_(String(q.email || '')),
  codeVerifier: (q) => codeVerifier_(String(q.email || ''), String(q.code || '')),
};
// avec clé
const ACTIONS = {
  moi: (u, q) => { if (u.statut === 'attente' && !lignes_('demandes').some(r => adr_(r[0]) === u.email)) noterDemande_(u.email, '', true); return moi_(u, q); },
  demande: (u, q) => (noterDemande_(u.email, String(q.classe || ''), true), moi_(u, {})),
  envoyer: (u, q) => envoyer_(u, Array.isArray(q.ev) ? q.ev : []),
  profDonnees: (u) => (prof_(u), profDonnees_(u)),
  profAutoriser: (u, q) => (prof_(u), autoriser_(u, String(q.adresse || ''), String(q.classe || ''), String(q.nom || '')), profDonnees_(u)),
  profRefuser: (u, q) => (prof_(u), refuser_(u, String(q.adresse || '')), profDonnees_(u)),
  profRetirer: (u, q) => (prof_(u), retirer_(String(q.adresse || '')), profDonnees_(u)),
  profClasse: (u, q) => (prof_(u), changerClasse_(String(q.adresse || ''), String(q.classe || '')), profDonnees_(u)),
};

/* =========================================================
   Clé signée (HMAC-SHA256)
   ========================================================= */
function secret_() { return PropertiesService.getScriptProperties().getProperty('SECRET'); }
function b64_(s) { return Utilities.base64EncodeWebSafe(s).replace(/=+$/, ''); }
function signer_(payload) {
  const p = b64_(JSON.stringify(payload));
  return p + '.' + b64_(Utilities.computeHmacSha256Signature(p, secret_()));
}
function erreurAuth_(msg) { const e = new Error(msg); e.auth = true; return e; }
function verifier_(cle) {
  const parts = String(cle || '').split('.');
  if (parts.length !== 2) throw erreurAuth_('Connecte-toi d’abord.');
  if (b64_(Utilities.computeHmacSha256Signature(parts[0], secret_())) !== parts[1]) throw erreurAuth_('Connexion non valide : reconnecte-toi.');
  const pad = parts[0] + '==='.slice((parts[0].length + 3) % 4);
  const p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(pad)).getDataAsString());
  if (!p.e || Date.now() > p.x) throw erreurAuth_('Ta connexion a expiré : reconnecte-toi.');
  return String(p.e).toLowerCase();
}

/* =========================================================
   Outils
   ========================================================= */
function onglet_(k) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLETS[k].nom); }
function lignes_(k) {
  const sh = onglet_(k), n = sh.getLastRow();
  return n < 2 ? [] : sh.getRange(2, 1, n - 1, ONGLETS[k].entetes.length).getValues();
}
const adr_ = s => String(s || '').trim().toLowerCase();
function cap_(s) { return String(s || '').split('-').map(p => p ? p.charAt(0).toUpperCase() + p.slice(1) : p).join('-'); }
function html_(s) { return String(s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function verrou_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }
const dateMs_ = d => d instanceof Date ? d.getTime() : (d ? new Date(d).getTime() || 0 : 0);
function emailValide_(e) { return /^[^\s@<>()",;:]+@[a-z0-9.-]+\.[a-z]{2,}$/.test(e); }
function nomDe_(email) {
  const a = lignes_('autorises').find(r => adr_(r[0]) === email);
  if (a && String(a[1]).trim()) { const p = String(a[1]).trim().split(/\s+/); return { nom: p[0], prenom: p.slice(1).join(' '), complet: String(a[1]).trim() }; }
  const parts = email.split('@')[0].split(/[._]/).filter(Boolean);
  if (parts.length > 1 && /^students?\d*$/.test(parts[parts.length - 1])) parts.pop();
  const prenom = cap_(parts[0] || ''), nom = parts.slice(1).map(cap_).join(' ');
  return { prenom, nom, complet: (nom ? nom + ' ' : '') + prenom };
}
function profDe_(email) { return lignes_('profs').find(r => adr_(r[0]) === email); }
function prof_(u) { if (u.role !== 'prof') throw new Error('Réservé aux professeurs.'); }
function listeClasses_() { return lignes_('classes').filter(r => String(r[0]).trim()).map(r => ({ classe: String(r[0]).trim(), prof: adr_(r[1]) })); }

/* Statut d'une adresse : prof, autorisée (adresse ou @domaine), refusée ou en attente */
function statut_(email) {
  if (profDe_(email)) return { role: 'prof', statut: 'ok', classe: '' };
  const dom = '@' + email.split('@')[1];
  const aut = lignes_('autorises');
  const a = aut.find(r => adr_(r[0]) === email) || aut.find(r => adr_(r[0]) === dom);
  if (a) return { role: 'eleve', statut: 'ok', classe: String(a[2] || '').trim() };
  const d = lignes_('demandes').find(r => adr_(r[0]) === email);
  return { role: 'eleve', statut: d && /refus/i.test(String(d[5])) ? 'refuse' : 'attente', classe: d ? String(d[2] || '').trim() : '' };
}
function noterDemande_(email, classe, verifiee) {
  return verrou_(() => {
    const sh = onglet_('demandes'), rows = lignes_('demandes');
    const i = rows.findIndex(r => adr_(r[0]) === email);
    if (i < 0) { const n = nomDe_(email); sh.appendRow([email, n.complet, classe, new Date(), !!verifiee, 'En attente', '', '']); return; }
    if (classe) sh.getRange(i + 2, 3).setValue(classe);
    if (verifiee && rows[i][4] !== true) sh.getRange(i + 2, 5).setValue(true);
  });
}

/* =========================================================
   Connexion par code reçu par e-mail
   ========================================================= */
function codeDemander_(brut) {
  const email = adr_(brut);
  if (!emailValide_(email)) throw new Error('Adresse e-mail non valide.');
  const cache = CacheService.getScriptCache();
  const st = statut_(email), ecole = email.split('@')[1] === DOMAINE;
  if (st.statut !== 'ok' && !ecole) {
    // adresse extérieure pas encore autorisée : on transmet la demande au professeur (identité non vérifiée)
    const g = Number(cache.get('demandes-ext') || 0);
    if (g < 30) { cache.put('demandes-ext', String(g + 1), 3600); if (!lignes_('demandes').some(r => adr_(r[0]) === email)) noterDemande_(email, '', false); }
    return { envoye: false, message: 'Cette adresse n’est pas encore autorisée. Ta demande est transmise au professeur : quand il l’aura acceptée, reviens demander ton code.' };
  }
  const nb = Number(cache.get('n:' + email) || 0);
  if (nb >= 3) throw new Error('Trop de codes demandés pour cette adresse. Réessaie dans un quart d’heure.');
  if (MailApp.getRemainingDailyQuota() < 1) throw new Error('Plus d’envoi possible aujourd’hui. Réessaie demain ou connecte-toi avec ton compte de l’école.');
  const code = String(Math.floor(100000 + Math.random() * 900000));
  cache.put('n:' + email, String(nb + 1), CODE_MINUTES * 60);
  cache.put('c:' + email, JSON.stringify({ h: b64_(Utilities.computeHmacSha256Signature(code, secret_())), essais: 0 }), CODE_MINUTES * 60);
  MailApp.sendEmail({
    to: email, name: 'Banc de câblage', subject: 'Ton code de connexion : ' + code,
    htmlBody: `<p>Bonjour,</p><p>Voici ton code pour te connecter au Banc de câblage :</p>
      <p style="font-size:28px;font-weight:bold;letter-spacing:4px">${code}</p>
      <p>Il est valable ${CODE_MINUTES} minutes. Si tu n’as rien demandé, ignore ce message.</p>`,
  });
  return { envoye: true };
}
function codeVerifier_(brut, code) {
  const email = adr_(brut), cache = CacheService.getScriptCache();
  const c = JSON.parse(cache.get('c:' + email) || 'null');
  if (!c) throw new Error('Code expiré : demande un nouveau code.');
  if (c.essais >= 5) { cache.remove('c:' + email); throw new Error('Trop d’essais : demande un nouveau code.'); }
  if (b64_(Utilities.computeHmacSha256Signature(String(code).replace(/\D/g, ''), secret_())) !== c.h) {
    c.essais++; cache.put('c:' + email, JSON.stringify(c), CODE_MINUTES * 60);
    throw new Error('Code incorrect.');
  }
  cache.remove('c:' + email);
  if (statut_(email).statut !== 'ok') noterDemande_(email, '', true);
  return { cle: signer_({ e: email, x: Date.now() + DUREE_CLE_JOURS * 864e5 }) };
}

/* =========================================================
   Élève
   ========================================================= */
function moi_(u) {
  const n = nomDe_(u.email);
  const res = { email: u.email, role: u.role, statut: u.statut, classe: u.classe, nom: n.nom, prenom: n.prenom, complet: n.complet, classes: listeClasses_().map(c => c.classe) };
  if (u.role === 'prof') {
    const p = profDe_(u.email); if (p && String(p[1]).trim()) res.complet = String(p[1]).trim();
    const cp = PropertiesService.getScriptProperties().getProperty('CODE_PROF'); if (cp) res.codeProf = cp;
  }
  return res;
}
function envoyer_(u, evs) {
  if (u.statut !== 'ok') { const e = new Error('Accès pas encore autorisé.'); e.attente = true; throw e; }
  if (evs.length > 200) evs = evs.slice(0, 200);
  return verrou_(() => {
    const jr = onglet_('journal'), n = jr.getLastRow();
    const vus = new Set(n < 2 ? [] : jr.getRange(2, 9, n - 1, 1).getValues().map(r => String(r[0])));
    const neufs = evs.filter(e => e && e.id && !vus.has(String(e.id)) && String(e.k || '').length < 80);
    const recus = evs.filter(e => e && e.id).map(e => String(e.id));
    if (!neufs.length) return { recus };
    const nm = nomDe_(u.email).complet;
    jr.getRange(n + 1, 1, neufs.length, 9).setValues(neufs.map(e => [new Date(Number(e.t) || Date.now()), u.email, String(e.k), String(e.titre || '').slice(0, 120), String(e.lvl || ''), String(e.type || ''),
      e.type === 'panne' ? 'Panne trouvée' : e.ok ? 'Réussi' : 'À corriger', String(e.detail || '').slice(0, 300), String(e.id)]));
    const rs = onglet_('resultats'), rows = lignes_('resultats');
    const idx = new Map(rows.map((r, i) => [adr_(r[0]) + '|' + r[3], i]));
    const ajout = [];
    neufs.sort((a, b) => (Number(a.t) || 0) - (Number(b.t) || 0)).forEach(e => {
      const cle = u.email + '|' + e.k, t = new Date(Number(e.t) || Date.now());
      let r, i = idx.get(cle);
      if (i === undefined) { r = [u.email, nm, u.classe, String(e.k), String(e.titre || ''), String(e.lvl || ''), String(e.type || ''), 0, false, '', t, '', t, '']; rows.push(r); i = rows.length - 1; idx.set(cle, i); ajout.push(i); }
      r = rows[i];
      r[1] = nm; r[2] = u.classe || r[2]; r[7] = (Number(r[7]) || 0) + 1; r[12] = t;
      if ((e.ok || e.type === 'panne') && r[8] !== true) { r[8] = true; r[9] = r[7]; r[11] = t; }
      if (e.type === 'panne' && e.note != null && (r[13] === '' || Number(e.note) > Number(r[13]))) r[13] = Number(e.note);
      r.changed = true;
    });
    rows.forEach((r, i) => { if (r.changed && !ajout.includes(i)) rs.getRange(i + 2, 1, 1, 14).setValues([r.slice(0, 14)]); });
    if (ajout.length) { const d = rs.getLastRow() + 1; rs.getRange(d, 1, ajout.length, 14).setValues(ajout.map(i => rows[i].slice(0, 14))); }
    return { recus };
  });
}

/* =========================================================
   Professeur : accès et suivi
   ========================================================= */
function profDonnees_(u) {
  const classes = listeClasses_();
  const mesClasses = classes.filter(c => c.prof === u.email).map(c => c.classe);
  const aut = lignes_('autorises').filter(r => String(r[0]).trim()).map(r => ({ adresse: adr_(r[0]), nom: String(r[1] || ''), classe: String(r[2] || ''), par: adr_(r[3]), le: dateMs_(r[4]) }));
  const dem = lignes_('demandes').filter(r => String(r[0]).trim() && !/accept/i.test(String(r[5])))
    .map(r => ({ adresse: adr_(r[0]), nom: String(r[1] || ''), classe: String(r[2] || ''), le: dateMs_(r[3]), verifiee: r[4] === true, statut: String(r[5] || '') }));
  const res = lignes_('resultats').filter(r => String(r[0]).trim()).map(r => ({ adresse: adr_(r[0]), nom: String(r[1] || ''), classe: String(r[2] || ''), k: String(r[3]), titre: String(r[4]), lvl: String(r[5]), type: String(r[6]),
    essais: Number(r[7]) || 0, reussi: r[8] === true, n: Number(r[9]) || 0, premier: dateMs_(r[10]), reussiLe: dateMs_(r[11]), dernier: dateMs_(r[12]), note: r[13] === '' ? null : Number(r[13]) }));
  return { moi: u.email, classes: classes.map(c => c.classe), mesClasses, autorises: aut, demandes: dem, resultats: res };
}
function autoriser_(u, brut, classe, nom) {
  const a = adr_(brut);
  if (!emailValide_(a) && !/^@[a-z0-9.-]+\.[a-z]{2,}$/.test(a)) throw new Error('Indique une adresse e-mail ou un domaine qui commence par @.');
  verrou_(() => {
    const sh = onglet_('autorises'), rows = lignes_('autorises'), i = rows.findIndex(r => adr_(r[0]) === a);
    const dm = lignes_('demandes').find(r => adr_(r[0]) === a);
    const nm = nom || (i >= 0 ? String(rows[i][1]) : '') || (dm && !/@/.test(String(dm[1])) ? String(dm[1]) : '');
    const cl = classe || (i >= 0 ? String(rows[i][2]) : '') || (dm ? String(dm[2] || '') : '');
    if (i < 0) sh.appendRow([a, nm, cl, u.email, new Date(), '']);
    else sh.getRange(i + 2, 2, 1, 2).setValues([[nm, cl]]);
    marquerDemande_(a, 'Acceptée', u.email);
  });
}
function refuser_(u, brut) { const a = adr_(brut); verrou_(() => marquerDemande_(a, 'Refusée', u.email)); }
function marquerDemande_(a, statut, par) {
  const sh = onglet_('demandes'), rows = lignes_('demandes'), i = rows.findIndex(r => adr_(r[0]) === a);
  if (i >= 0) sh.getRange(i + 2, 6, 1, 3).setValues([[statut, par, new Date()]]);
}
function retirer_(brut) {
  const a = adr_(brut);
  verrou_(() => {
    const sh = onglet_('autorises'), rows = lignes_('autorises');
    for (let i = rows.length - 1; i >= 0; i--) if (adr_(rows[i][0]) === a) sh.deleteRow(i + 2);
    marquerDemande_(a, 'Accès retiré', '');
  });
}
function changerClasse_(brut, classe) {
  const a = adr_(brut);
  verrou_(() => {
    const sh = onglet_('autorises'), rows = lignes_('autorises'), i = rows.findIndex(r => adr_(r[0]) === a);
    if (i >= 0) sh.getRange(i + 2, 3).setValue(classe);
  });
}
