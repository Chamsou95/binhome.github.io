/**
 * PORTRAIT OLFACTIF — réception des questionnaires
 * ------------------------------------------------
 * À coller dans Extensions > Apps Script d'une feuille Google Sheets,
 * puis à déployer en « Application web » (voir les instructions en bas).
 */

/* ============ RÉGLAGES ============ */
const NOTIFIER = "contact@binhome.fr";   // e-mail prévenu à chaque portrait (vide = pas d'e-mail)
const ONGLET   = "Portraits";            // nom de l'onglet où les lignes s'ajoutent
/* ================================== */

const COLONNES = [
  "Reçu le","Réf","Prénom","E-mail","Offre","Pour qui","Occasion","Occasion (détail)","Niveau",
  "Le moment","Souvenir marquant","Émotions","Perception","Pourquoi maintenant","Lassé par","Mot-clé",
  "Images choisies","Images écartées",
  "Odeurs aimées","Odeurs rejetées","Odeurs indifférentes","Odeur préférée","Son souvenir",
  "Frais→Chaud","Léger→Enveloppant","Sec→Sucré","Propre→Charnel","Naturel→Abstrait","Discret→Sillage","Classique→Surprenant",
  "Tenue","Codes",
  "Parfums cités","Aime sur les autres","Déteste",
  "Quand","Saisons","Environnement","Zones","Réaction passée","Réaction (détail)","Note précieuse",
  "Restitution","Mot libre","Consentement","JSON complet"
];

const LIEUX = {
  sousbois:"Sous-bois après l'averse", crique:"Crique méditerranéenne", souk:"Souk aux épices",
  biblio:"Bibliothèque ancienne", grange:"Grange en fin d'été", ville:"Ville après la pluie, la nuit",
  boulangerie:"Boulangerie au petit matin", jardin:"Jardin en fleurs au crépuscule",
  lin:"Lin froissé", laine:"Laine douce", velours:"Velours", cuir:"Cuir patiné", bois:"Bois brut", pierre:"Pierre mouillée",
  aube:"L'aube", midi:"Plein midi", dore:"Fin d'après-midi dorée", nuit:"La nuit",
  blancs:"Blancs et gris clairs", verts:"Verts profonds", ocres:"Ocres et terracotta",
  roses:"Roses poudrés", bleus:"Bleus froids", sombres:"Noirs et bruns"
};
const SERIES = ["lieu","matiere","moment","couleurs"];
const CURSEURS = ["temperature","densite","douceur","peau","nature","presence","audace"];

/* ---------- Réception ---------- */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const r = data.reponses || {};
    const feuille = feuilleCible();
    ecrire(feuille, ligne(data, r));
    colorerDerniereLigne(feuille);
    if (NOTIFIER) prevenir(data, r);
    return reponse({ ok: true, ref: data.id });
  } catch (err) {
    journaliserErreur(err, e);
    return reponse({ ok: false, erreur: String(err) });
  }
}

function doGet() {
  return reponse({ ok: true, message: "Réception des portraits opérationnelle." });
}

/* ---------- Construction de la ligne ---------- */
function ligne(data, r) {
  const cartes = r.cartes || {};
  const parCategorie = v => Object.keys(cartes).filter(k => cartes[k] === v).join(", ");

  const images = [], ecartees = [];
  SERIES.forEach(s => {
    (r["img_" + s] || []).forEach(id => images.push(LIEUX[id] || id));
    if (r["rej_" + s]) ecartees.push(LIEUX[r["rej_" + s]] || r["rej_" + s]);
  });

  const parfums = (r.parfums || [])
    .filter(p => p && p.nom)
    .map(p => p.nom + (p.verdict ? " (" + p.verdict + ")" : "") + (p.raisons && p.raisons.length ? " — " + p.raisons.join(", ") : ""))
    .join("\n");

  const OFFRES = { decouverte:"Découverte", signature:"Signature", exclusive:"Exclusive", indecis:"À définir" };

  const l = [
    new Date(data.envoye_le || Date.now()), data.id || "",
    r.prenom || "", r.email || "", OFFRES[r.offre] || r.offre || "", r.pour_qui || "", r.occasion || "", r.occasion_autre || "", r.niveau || "",
    r.moment || "", r.souvenir_fort || "", (r.emotions || []).join(", "), r.perception || "", r.pourquoi || "", r.lasse_quoi || "", r.mot || "",
    images.join(", "), ecartees.join(", "),
    parCategorie("aime"), parCategorie("non"), parCategorie("neutre"), r.carte_pref || "", r.carte_souvenir || ""
  ];
  CURSEURS.forEach(c => l.push(r["c_" + c] === undefined ? "" : r["c_" + c]));
  l.push(
    r.tenue || "", r.codes || "",
    parfums, r.aime_autre || "", r.deteste || "",
    (r.quand || []).join(", "), (r.saisons || []).join(", "), r.environnement || "", (r.zones || []).join(", "),
    r.reaction || "", r.reaction_quoi || "", r.precieuse || "",
    r.restitution || "", r.libre || "", r.ok_donnees ? "oui" : "non",
    JSON.stringify(data)
  );
  return l;
}

/* ---------- Écriture tolérante aux anciens en-têtes ---------- */
function ecrire(f, valeurs) {
  const entetes = f.getRange(1, 1, 1, Math.max(f.getLastColumn(), 1)).getValues()[0].map(String);
  // Toute colonne attendue et absente est ajoutée à la fin de l'en-tête.
  COLONNES.forEach(nom => {
    if (entetes.indexOf(nom) === -1) {
      entetes.push(nom);
      const c = entetes.length;
      f.getRange(1, c).setValue(nom).setFontWeight("bold").setBackground("#23262B").setFontColor("#FFFFFF").setWrap(true);
    }
  });
  const sortie = entetes.map(nom => {
    const i = COLONNES.indexOf(nom);
    return i === -1 ? "" : valeurs[i];
  });
  f.appendRow(sortie);
}

/* ---------- Feuille ---------- */
function feuilleCible() {
  const classeur = SpreadsheetApp.getActiveSpreadsheet();
  let f = classeur.getSheetByName(ONGLET);
  if (!f) f = classeur.insertSheet(ONGLET);
  if (f.getLastRow() === 0) {
    f.appendRow(COLONNES);
    const entete = f.getRange(1, 1, 1, COLONNES.length);
    entete.setFontWeight("bold").setBackground("#23262B").setFontColor("#FFFFFF").setWrap(true);
    f.setFrozenRows(1);
    f.setColumnWidth(1, 140);
    f.setColumnWidth(COLONNES.length, 90);       // JSON complet, volontairement étroit
    for (let c = 9; c <= 22; c++) f.setColumnWidth(c, 220);
  }
  return f;
}

function colorerDerniereLigne(f) {
  const n = f.getLastRow();
  f.getRange(n, 1, 1, COLONNES.length).setVerticalAlignment("top").setWrap(true);
  const entetes = f.getRange(1, 1, 1, f.getLastColumn()).getValues()[0].map(String);
  const reaction = entetes.indexOf("Réaction passée") + 1;
  if (!reaction) return;
  if (f.getRange(n, reaction).getValue() === "oui") {
    f.getRange(n, reaction, 1, 2).setBackground("#F4E3E6");  // signalé : à lire avant de composer
  }
}

/* ---------- Notification ---------- */
function prevenir(data, r) {
  const cartes = r.cartes || {};
  const aimees = Object.keys(cartes).filter(k => cartes[k] === "aime");
  const rejetees = Object.keys(cartes).filter(k => cartes[k] === "non");
  const lien = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const corps = [
    "Nouveau portrait olfactif.",
    "",
    "Prénom : " + (r.prenom || "—"),
    "E-mail : " + (r.email || "—"),
    "Formule : " + (r.offre || "non précisée"),
    "Pour : " + (r.pour_qui === "offrir" ? "offrir" : "lui-même") + " · " + (r.occasion || "—"),
    "",
    "Le moment : " + (r.moment || "—"),
    "Émotions : " + (r.emotions || []).join(", "),
    "Mot-clé : " + (r.mot || "—"),
    "",
    "Odeurs aimées : " + aimees.join(", "),
    "Odeurs rejetées : " + rejetees.join(", "),
    "",
    "Tenue : " + (r.tenue || "—") + " · Codes : " + (r.codes || "—"),
    "Réaction passée signalée : " + (r.reaction === "oui" ? "OUI — " + (r.reaction_quoi || "sans détail") : "non"),
    "",
    "Tout le détail dans la feuille : " + lien
  ].join("\n");
  MailApp.sendEmail({
    to: NOTIFIER,
    subject: "Portrait olfactif — " + (r.prenom || "sans prénom"),
    body: corps
  });
}

/* ---------- Utilitaires ---------- */
function reponse(objet) {
  return ContentService.createTextOutput(JSON.stringify(objet)).setMimeType(ContentService.MimeType.JSON);
}

function journaliserErreur(err, e) {
  try {
    const classeur = SpreadsheetApp.getActiveSpreadsheet();
    let f = classeur.getSheetByName("Erreurs") || classeur.insertSheet("Erreurs");
    f.appendRow([new Date(), String(err), e && e.postData ? e.postData.contents : ""]);
  } catch (_) {}
}

/* ---------- Test depuis l'éditeur ---------- */
function testerAvecUnFauxPortrait() {
  const faux = {
    id: "test" + Date.now(), envoye_le: new Date().toISOString(),
    reponses: {
      prenom: "Camille", email: "camille@exemple.fr", offre: "signature", pour_qui: "moi", occasion: "sansraison", niveau: "amateur",
      moment: "le premier café sur le balcon", emotions: ["apaisé", "lumineux"], perception: "douceur",
      pourquoi: "cadeau_soi", mot: "refuge",
      img_lieu: ["boulangerie"], rej_lieu: "ville", img_matiere: ["laine"], img_moment: ["aube"], img_couleurs: ["ocres"],
      cartes: { "Pain chaud": "aime", "Miel": "aime", "Fleur d'oranger": "aime", "Cuir": "non" },
      carte_pref: "Pain chaud", carte_souvenir: "la boulangerie de mon grand-père",
      c_temperature: 7, c_densite: 6, c_douceur: 4, c_peau: 5, c_nature: 4, c_presence: 3, c_audace: 6,
      tenue: "journee", codes: "entre",
      parfums: [{ nom: "Un parfum connu", verdict: "adore", raisons: ["il tient bien"] }],
      deteste: "les grands sucrés",
      quand: ["quotidien", "travail"], saisons: ["toute"], environnement: "openspace", zones: ["poignets", "cou"],
      reaction: "non", precieuse: "surprends", restitution: "histoire", ok_donnees: true
    }
  };
  const res = doPost({ postData: { contents: JSON.stringify(faux) } });
  Logger.log(res.getContent());
}

/* ============================================================
   INSTALLATION
   ------------------------------------------------------------
   1. Crée une feuille Google Sheets, nomme-la « Portraits olfactifs ».
   2. Extensions > Apps Script. Efface le contenu, colle tout ce fichier.
   3. Change NOTIFIER en haut pour ton adresse e-mail.
   4. Enregistre, puis lance testerAvecUnFauxPortrait (menu Exécuter).
      Google demande une autorisation : accepte-la. Tu dois voir
      apparaître une ligne « Camille » dans la feuille, et recevoir un e-mail.
   5. Déployer > Nouveau déploiement > type « Application web ».
      Exécuter en tant que : moi.
      Qui a accès : tout le monde.  (indispensable : le site poste sans compte Google)
   6. Copie l'URL de déploiement (elle finit par /exec) et colle-la dans
      portrait-olfactif.html, dans CONFIG.endpoint.
   7. Remplis un portrait depuis le site : une ligne doit s'ajouter.

   Mise à jour : si ta feuille existe déjà avec d'anciens en-têtes, le script
   ajoute lui-même les colonnes manquantes (comme « Offre ») à la fin de la
   ligne d'en-tête, et remplit chaque valeur d'après le NOM de la colonne.
   Aucune ligne déjà enregistrée n'est déplacée.

   À savoir : après chaque modification du script, il faut
   Déployer > Gérer les déploiements > modifier > nouvelle version,
   sinon l'ancienne version continue de tourner.
   ============================================================ */
