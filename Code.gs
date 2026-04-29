/**
 * PELICHET NLC SA — Backend Google Apps Script v2.0
 *
 * Refonte : validation serveur, LockService, corrections de bugs,
 * configuration centralisée, logging structuré.
 */

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
  TEMPLATE_PELICHET_ID: '141tclld00jgVrxuQ4b4T738LDkk2A19j21mfmJTrDDk',
  TEMPLATE_AUTRE_ID: '1QWDp-ACk1dL7bXF2fq5VQRbPW4jxwXtUmq8hA1z4cD0',
  TEMPLATE_RESA_ID: '1U2hICEGuhzv9acMZ6MeeykIB41Lx3H0JhVEKL_8sABA',
  FOLDER_ID: '1MP1I55oDhisTnm4zFJmki1Wap3fUff6U',
  SIGNATURES_FOLDER_ID: '15EmN3RiLKjH5i43zae6BvRlA3LKniTBF',
  SIGNATURES_FOLDER_NAME: 'Signatures',
  SPREADSHEET_ID: '1AUfeykbUZ07SG-WkqVuxWJY6plH43EZp0tJgrp_gwYM',
  COLOR_PELICHET: '#D32F2F',
  TVA_RATE: 0.081,
  RPLP_RATE: 0.005,
  TIMEZONE: 'GMT+1',
  LOCK_TIMEOUT_MS: 30000,
  SHEET_NAME: 'Suivi_Devis',  // Prefixe : Suivi_Devis_{USER}
  USERS_SHEET: 'Users',

  // Mapping ressources -> postes comptables pour la ventilation
  VENTILATION: {
    personnel: {
      poste: 'FR-1100-25',
      ventil: 'LOC2',
      libelle: "Main d'oeuvre demenagement local"
    },
    personnel_demi: {
      poste: 'FR-1100-02',
      ventil: 'LOC2',
      libelle: 'Portage (0 m) (1/2 journee homme)'
    },
    vehicules: {
      poste: 'FR-1100-16',
      ventil: 'LOC5',
      libelle: 'Traction Domicile-Domicile'
    },
    engins: {
      _default: { poste: 'FR-5220-01', ventil: 'LOC9', libelle: 'Monte meuble ou echelle electrique' },
      'Monte-meuble': { poste: 'FR-5220-01', ventil: 'LOC9', libelle: 'Monte meuble ou echelle electrique' },
      'Grue mobile': { poste: 'FR-5220-01', ventil: 'LOC9', libelle: 'Grue mobile' },
      'Chariot elevateur': { poste: 'FR-5220-01', ventil: 'LOC9', libelle: 'Chariot elevateur' }
    },
    materiel_emballage: {
      poste: 'FR-6800-08',
      ventil: 'LOC1',
      libelle: 'Emballage Standard terrestre'
    },
    materiel_specifique: {
      poste: 'FR-6800-03',
      ventil: 'LOC9',
      libelle: 'Emballages specifiques (Crate, caisse)'
    },
    gasoil: {
      poste: 'FR-4300-01',
      ventil: 'LOC5',
      libelle: 'Gasoile et/ou peage'
    },
    km_supplement: {
      poste: 'FR-4300-01',
      ventil: 'LOC5',
      libelle: 'Supplement kilometrique'
    },
    soustraitance: {
      poste: 'FR-1100-04',
      ventil: 'LOC7',
      libelle: 'Sous-traitance diverses au chargement'
    },
    stationnement: {
      poste: 'FR-1250-01',
      ventil: 'LOC9',
      libelle: 'Autorisation de stationnement'
    },
    panier_repas: {
      poste: 'FR-1100-11',
      ventil: 'LOCG',
      libelle: 'Panier repas/Hotel'
    },
    etages: {
      poste: 'FR-1100-03',
      ventil: 'LOC3',
      libelle: 'Etages'
    },
    manutention_gm: {
      poste: 'FR-1100-08',
      ventil: 'LOCB',
      libelle: 'Manutention entree/sortie Garde Meuble'
    },
    autres_sans_facture: {
      poste: 'FR-1100-20',
      ventil: 'LOC8',
      libelle: 'Autres charges (sans factures)'
    },
    autres_avec_facture: {
      poste: 'FR-1100-23',
      ventil: 'LOC9',
      libelle: 'Autres charges (avec factures)'
    },
    livraison_1: {
      poste: 'FR-1100-13',
      ventil: 'LOCI',
      libelle: 'Autres charges de livraison'
    },
    livraison_2: {
      poste: 'FR-1100-14',
      ventil: 'LOCI',
      libelle: 'Autres charges de livraison'
    }
  }
};

// ============================================
// UTILITAIRES
// ============================================

/**
 * Formate un nombre en CHF
 */
function fmt(v) {
  const num = parseFloat(v);
  if (isNaN(num)) return '0.00 CHF';
  return num.toLocaleString('fr-CH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' CHF';
}

/**
 * Calcule le kilométrage du trajet complet :
 * Pelichet Vernier → Départ → Arrivée → Pelichet Vernier
 * Retourne { totalKm, forfaitKm, excessKm, legs[] }
 */
function calculerKilometrage(adresseDepart, adresseArrivee) {
  var PELICHET = 'Chemin Francois-Lehmann 12, 1218 Le Grand-Saconnex, Suisse';
  var FORFAIT_KM = 50; // 25 aller + 25 retour, canton de Genève

  var directions = Maps.newDirectionFinder()
    .setOrigin(PELICHET)
    .setDestination(PELICHET)
    .addWaypoint(adresseDepart)
    .addWaypoint(adresseArrivee)
    .setMode(Maps.DirectionFinder.Mode.DRIVING)
    .setLanguage('fr')
    .getDirections();

  if (directions.status !== 'OK') {
    Logger.log('Maps Directions error: ' + directions.status);
    return { error: 'Impossible de calculer l\'itineraire (' + directions.status + ')' };
  }

  var totalMeters = 0;
  var legs = [];
  var route = directions.routes[0];

  for (var i = 0; i < route.legs.length; i++) {
    var leg = route.legs[i];
    totalMeters += leg.distance.value;
    legs.push({
      from: leg.start_address,
      to: leg.end_address,
      km: Math.round(leg.distance.value / 1000),
      duration: leg.duration.text
    });
  }

  var totalKm = Math.round(totalMeters / 1000);
  var excessKm = Math.max(0, totalKm - FORFAIT_KM);

  return {
    totalKm: totalKm,
    forfaitKm: FORFAIT_KM,
    excessKm: excessKm,
    legs: legs
  };
}

/**
 * Recherche d'entreprises / adresses suisses via Nominatim (OpenStreetMap)
 * API gratuite, sans clé, contient les entreprises suisses
 * @param {string} query - Terme de recherche (nom d'entreprise)
 * @returns {Array} Liste de résultats {name, street, zip, city}
 */
function searchSwissCompany(query) {
  var url = 'https://nominatim.openstreetmap.org/search'
    + '?q=' + encodeURIComponent(query)
    + '&countrycodes=ch'
    + '&format=json'
    + '&addressdetails=1'
    + '&limit=8';

  var options = {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'PelichetLogistique/1.0' }
  };

  var response = UrlFetchApp.fetch(url, options);
  var code = response.getResponseCode();
  if (code !== 200) {
    Logger.log('Nominatim API error: ' + code);
    return [];
  }

  var data = JSON.parse(response.getContentText());
  if (!Array.isArray(data)) return [];

  // Dédupliquer par nom
  var seen = {};
  var results = [];

  data.forEach(function(item) {
    var addr = item.address || {};
    var name = item.name || '';
    var street = addr.road || '';
    if (addr.house_number) street += ' ' + addr.house_number;
    var zip = addr.postcode || '';
    var city = addr.city || addr.town || addr.village || addr.municipality || '';

    // Clé de déduplication
    var key = (name + '|' + zip).toLowerCase();
    if (seen[key]) return;
    seen[key] = true;

    results.push({
      name: name,
      street: street,
      zip: zip,
      city: city
    });
  });

  return results;
}

/**
 * Calcule la date d'intervention prévue à partir des RDV des postes.
 * Si un seul poste avec RDV → affiche cette date
 * Si plusieurs postes/RDV → affiche la plage "du dd.MM.yyyy au dd.MM.yyyy"
 * Fallback sur data.datePrevue si présent, sinon chaîne vide
 */
function calculerDatePrevue(data) {
  var dates = [];

  function parseDate(v) {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    // Format ISO "YYYY-MM-DD" du champ <input type="date">
    var isoMatch = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      // Force midi local pour eviter decalage timezone
      return new Date(parseInt(isoMatch[1]), parseInt(isoMatch[2]) - 1, parseInt(isoMatch[3]), 12, 0, 0);
    }
    // Format "dd.MM.yyyy" ou "dd/MM/yyyy"
    var frMatch = s.match(/^(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})/);
    if (frMatch) {
      return new Date(parseInt(frMatch[3]), parseInt(frMatch[2]) - 1, parseInt(frMatch[1]), 12, 0, 0);
    }
    // Fallback natif
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }

  safeArray(data.postes).forEach(function(p) {
    safeArray(p.rdvs).forEach(function(r) {
      var d = parseDate(r.date);
      if (d) dates.push(d);
    });
  });

  Logger.log('calculerDatePrevue: ' + dates.length + ' dates trouvees');

  if (dates.length === 0) {
    // Fallback manuel si champ existant
    var fallback = parseDate(data.datePrevue);
    if (fallback) return Utilities.formatDate(fallback, CONFIG.TIMEZONE, 'dd.MM.yyyy');
    return safe(data.datePrevue);
  }

  // Trier chronologiquement
  dates.sort(function(a, b) { return a - b; });
  var premiere = Utilities.formatDate(dates[0], CONFIG.TIMEZONE, 'dd.MM.yyyy');

  if (dates.length === 1) return premiere;

  var derniere = Utilities.formatDate(dates[dates.length - 1], CONFIG.TIMEZONE, 'dd.MM.yyyy');
  if (premiere === derniere) return premiere;

  return 'du ' + premiere + ' au ' + derniere;
}

/**
 * Abréviations officielles du personnel
 */
function transformerPersonnel(liste) {
  if (!Array.isArray(liste)) return [];
  return liste.map(item =>
    String(item)
      .replace(/Manutentionnaire du lourd/gi, 'HL')
      .replace(/Manutentionnaire/gi, 'H')
      .replace(/Chauffeur-Livreur/gi, 'CHAUF-LIV')
      .replace(/CHAUF-LIVREUR/gi, 'CHAUF-LIV')
      .replace(/Chauffeur/gi, 'CHAUF')
  );
}

/**
 * Calcule la date de Paques (algorithme de Meeus/Jones/Butcher)
 */
function getEasterDate(year) {
  var a = year % 19;
  var b = Math.floor(year / 100);
  var c = year % 100;
  var d = Math.floor(b / 4);
  var e = b % 4;
  var f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4);
  var k = c % 4;
  var l = (32 + 2 * e + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var month = Math.floor((h + l - 7 * m + 114) / 31);
  var day = ((h + l - 7 * m + 114) % 31) + 1;
  return new Date(year, month - 1, day);
}

/**
 * Retourne la liste des jours feries suisses (Geneve) pour une annee donnee
 * Inclut : Nouvel An, 2 janvier, Vendredi Saint, Lundi de Paques,
 * Ascension, Lundi de Pentecote, 1er mai, 1er aout, Jeune genevois,
 * Noel, 31 decembre
 */
function getJoursFeries(year) {
  var easter = getEasterDate(year);
  var easterMs = easter.getTime();
  var day = 86400000; // ms par jour

  var feries = [
    new Date(year, 0, 1),   // Nouvel An
    new Date(year, 0, 2),   // 2 janvier (Geneve)
    new Date(easterMs - 2 * day),  // Vendredi Saint
    new Date(easterMs + 1 * day),  // Lundi de Paques
    new Date(easterMs + 39 * day), // Ascension (jeudi)
    new Date(easterMs + 50 * day), // Lundi de Pentecote
    new Date(year, 4, 1),   // 1er mai
    new Date(year, 7, 1),   // Fete nationale (1er aout)
    // Jeune genevois : jeudi apres le 1er dimanche de septembre
    (function() {
      var d = new Date(year, 8, 1); // 1er sept
      while (d.getDay() !== 0) d.setDate(d.getDate() + 1); // 1er dimanche
      d.setDate(d.getDate() + 4); // jeudi suivant
      return d;
    })(),
    new Date(year, 11, 25), // Noel
    new Date(year, 11, 31)  // 31 decembre (Geneve)
  ];

  return feries;
}

/**
 * Verifie si une date est un jour ferie
 */
function isJourFerie(date, feriesList) {
  var d = date.getFullYear() * 10000 + (date.getMonth() + 1) * 100 + date.getDate();
  for (var i = 0; i < feriesList.length; i++) {
    var f = feriesList[i].getFullYear() * 10000 + (feriesList[i].getMonth() + 1) * 100 + feriesList[i].getDate();
    if (d === f) return true;
  }
  return false;
}

/**
 * Verifie si une date est un jour ouvre (ni weekend, ni ferie)
 */
function isJourOuvre(date) {
  if (date.getDay() === 0 || date.getDay() === 6) return false;
  var feries = getJoursFeries(date.getFullYear());
  return !isJourFerie(date, feries);
}

/**
 * Retourne le prochain jour ouvre apres une date donnee
 */
function getNextWorkingDay(date) {
  var next = new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1, 12, 0, 0);
  while (!isJourOuvre(next)) {
    next = new Date(next.getFullYear(), next.getMonth(), next.getDate() + 1, 12, 0, 0);
  }
  return next;
}

/**
 * Genere la liste des jours ouvres a partir d'une date de debut pour N jours
 * Ex: getJoursOuvres(date, 10) -> tableau de 10 dates de jours ouvres
 */
function getJoursOuvres(dateDebut, nbJours) {
  var result = [];
  // Normaliser a midi pour eviter les decalages de timezone
  var current = new Date(dateDebut.getFullYear(), dateDebut.getMonth(), dateDebut.getDate(), 12, 0, 0);

  // Si la date de debut est ouvree, la compter
  if (isJourOuvre(current)) {
    result.push(new Date(current.getTime()));
  }

  while (result.length < Math.ceil(nbJours)) {
    current = getNextWorkingDay(current);
    result.push(new Date(current.getTime()));
  }

  return result;
}

/**
 * Ouvre le spreadsheet de suivi
 */
function getSs() {
  try {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  } catch (e) {
    try {
      return SpreadsheetApp.getActiveSpreadsheet();
    } catch (err) {
      Logger.log('ERREUR: Impossible d\'ouvrir le spreadsheet: ' + err);
      return null;
    }
  }
}

/**
 * Safe string: empêche les valeurs null/undefined
 */
function safe(val, fallback) {
  if (val === null || val === undefined || val === '') return fallback || '';
  return String(val);
}

/**
 * Safe array: garantit un tableau
 */
function safeArray(val) {
  return Array.isArray(val) ? val : [];
}

/**
 * Date du jour formatée
 */
function dateJour() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy');
}

// ============================================
// ITEMS PERSONNALISES (onglet Custom_Items)
// ============================================
var CUSTOM_SHEET_NAME = 'Custom_Items';

/**
 * Retourne tous les items personnalises { personnel: [...], vehicules: [...], ... }
 */
function getCustomItems() {
  var ss = getSs();
  if (!ss) return {};
  var sheet = ss.getSheetByName(CUSTOM_SHEET_NAME);
  if (!sheet) return {};

  var values = sheet.getDataRange().getValues();
  var result = {};
  // Format : colonne A = type, colonne B = valeur
  for (var i = 1; i < values.length; i++) {
    var type = String(values[i][0] || '').trim();
    var val = String(values[i][1] || '').trim();
    if (!type || !val) continue;
    if (!result[type]) result[type] = [];
    if (result[type].indexOf(val) === -1) result[type].push(val);
  }
  return result;
}

/**
 * Ajoute un item personnalise (type + valeur)
 */
function addCustomItem(type, value) {
  var ss = getSs();
  if (!ss) return false;
  var sheet = ss.getSheetByName(CUSTOM_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CUSTOM_SHEET_NAME);
    sheet.appendRow(['Type', 'Valeur', 'Date']);
    // Style en-tete
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');
  }

  // Verifier doublon
  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (String(existing[i][0]).trim() === type && String(existing[i][1]).trim() === value) {
      return false; // Deja existant
    }
  }

  sheet.appendRow([type, value, new Date()]);
  return true;
}

// ============================================
// GESTION UTILISATEURS (onglet Users)
// ============================================

/**
 * Retourne la liste des utilisateurs
 * Format : [{ id, nom, prenom, telephone, email, role }]
 */
function getUsers() {
  var ss = getSs();
  if (!ss) return [];
  var sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) {
    sheet = creerOngletUsers(ss);
  }

  var values = sheet.getDataRange().getValues();
  var users = [];
  // En-tetes : ID | Nom | Prenom | Telephone | Email | Role | Titre | SignatureId
  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][0] || '').trim();
    if (!id) continue;
    users.push({
      id: id,
      nom: String(values[i][1] || '').trim(),
      prenom: String(values[i][2] || '').trim(),
      telephone: String(values[i][3] || '').trim(),
      email: String(values[i][4] || '').trim(),
      role: String(values[i][5] || 'vendeur').trim(),
      titre: String(values[i][6] || '').trim(),
      signatureId: String(values[i][7] || '').trim()
    });
  }
  return users;
}

/**
 * Retourne un utilisateur par son ID
 */
function getUserById(userId) {
  var users = getUsers();
  for (var i = 0; i < users.length; i++) {
    if (users[i].id === userId) return users[i];
  }
  return null;
}

/**
 * Retourne le nom de l'onglet de suivi pour un utilisateur
 */
function getUserSheetName(userId) {
  if (!userId) return CONFIG.SHEET_NAME;
  return CONFIG.SHEET_NAME + '_' + userId.toUpperCase();
}

/**
 * S'assure que l'onglet de suivi d'un utilisateur existe (le cree si besoin)
 */
function ensureUserSheet(userId) {
  var ss = getSs();
  if (!ss) return null;
  var sheetName = getUserSheetName(userId);
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) return sheet;

  // Creer l'onglet avec les en-tetes
  sheet = ss.insertSheet(sheetName);
  var headers = ['Date', 'Ref', 'Societe', 'Email', 'Statut Email', 'Client',
                 'Facturation', 'Contact', 'Depart', 'Arrivee', 'Prevu',
                 'Detail Tarifie', 'HT Global', 'TTC Global', 'Lien', 'JSON'];
  sheet.appendRow(headers);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');

  // Figer la premiere ligne
  sheet.setFrozenRows(1);

  // Largeur des colonnes
  sheet.setColumnWidth(1, 130); // Date
  sheet.setColumnWidth(2, 160); // Ref
  sheet.setColumnWidth(6, 150); // Client
  sheet.setColumnWidth(13, 100); // HT
  sheet.setColumnWidth(14, 100); // TTC

  Logger.log('Onglet cree: ' + sheetName);
  return sheet;
}

/**
 * Ajoute un nouvel utilisateur dans l'onglet Users
 * Retourne l'objet user cree ou null si doublon
 */
function addUser(userData) {
  var ss = getSs();
  if (!ss) return null;
  var sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) sheet = creerOngletUsers(ss);

  var nom = String(userData.nom || '').trim().toUpperCase();
  var prenom = String(userData.prenom || '').trim();
  var telephone = String(userData.telephone || '').trim();
  var email = String(userData.email || '').trim();
  var role = String(userData.role || 'vendeur').trim();
  var titre = String(userData.titre || '').trim();
  var signatureBase64 = String(userData.signatureBase64 || '').trim();

  if (!nom || !prenom) return null;

  // Generer un ID unique : PRENOM (sans accents, majuscule, sans espaces)
  var id = prenom.toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]/g, '');

  // Verifier doublon d'ID
  var values = sheet.getDataRange().getValues();
  var suffix = '';
  var candidateId = id;
  var attempts = 0;
  while (attempts < 20) {
    var duplicate = false;
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim().toUpperCase() === candidateId) {
        duplicate = true;
        break;
      }
    }
    if (!duplicate) break;
    attempts++;
    candidateId = id + attempts;
  }
  id = candidateId;

  // Upload signature dans Google Drive si fournie
  var signatureId = '';
  if (signatureBase64) {
    try {
      signatureId = uploadSignature(id, signatureBase64);
    } catch (err) {
      Logger.log('Erreur upload signature: ' + err);
    }
  }

  // Ajouter la ligne SANS telephone (eviter #ERROR! sur +41...)
  sheet.appendRow([id, nom, prenom, '', email, role, titre, signatureId]);

  // Ecrire le telephone en forcant le format texte sur la cellule
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 4).setNumberFormat('@').setValue(telephone);

  // Creer l'onglet de suivi pour ce user
  ensureUserSheet(id);

  Logger.log('Utilisateur cree: ' + id + ' - ' + prenom + ' ' + nom + ' (signature: ' + signatureId + ')');

  return {
    id: id,
    nom: nom,
    prenom: prenom,
    telephone: telephone,
    email: email,
    role: role,
    titre: titre,
    signatureId: signatureId
  };
}

/**
 * Upload une image de signature dans Google Drive
 * signatureBase64 = "data:image/png;base64,iVBOR..." ou juste le base64
 * Retourne l'ID du fichier Drive
 */
function uploadSignature(userId, signatureBase64) {
  // Dossier dédié aux signatures (indépendant du dossier principal)
  var sigFolder;
  try {
    sigFolder = DriveApp.getFolderById(CONFIG.SIGNATURES_FOLDER_ID);
  } catch (e) {
    // Fallback : sous-dossier dans le dossier principal
    Logger.log('SIGNATURES_FOLDER_ID introuvable, fallback sous-dossier: ' + e);
    var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
    var sigFolders = folder.getFoldersByName(CONFIG.SIGNATURES_FOLDER_NAME);
    sigFolder = sigFolders.hasNext() ? sigFolders.next() : folder.createFolder(CONFIG.SIGNATURES_FOLDER_NAME);
  }

  // Extraire le type MIME et les donnees
  var mimeType = 'image/png';
  var base64Data = signatureBase64;
  if (signatureBase64.indexOf(',') !== -1) {
    var parts = signatureBase64.split(',');
    var header = parts[0]; // "data:image/png;base64"
    base64Data = parts[1];
    var mimeMatch = header.match(/data:([^;]+)/);
    if (mimeMatch) mimeType = mimeMatch[1];
  }

  // Decoder et creer le fichier
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType, 'signature_' + userId + '.png');
  var file = sigFolder.createFile(blob);

  // Rendre le fichier accessible (pour insertion dans les docs)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log('Signature uploadee: ' + file.getId() + ' pour ' + userId);
  return file.getId();
}

/**
 * Met a jour la signature d'un utilisateur existant
 */
function updateUserSignature(userId, signatureBase64) {
  var ss = getSs();
  if (!ss) return null;
  var sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return null;

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === userId.toUpperCase()) {
      // Upload la nouvelle signature
      var signatureId = uploadSignature(userId, signatureBase64);

      // Mettre a jour la colonne H (signatureId)
      sheet.getRange(i + 1, 8).setValue(signatureId);

      return signatureId;
    }
  }
  return null;
}

/**
 * Met a jour le titre d'un utilisateur existant
 */
function updateUserProfile(userId, titre) {
  var ss = getSs();
  if (!ss) return false;
  var sheet = ss.getSheetByName(CONFIG.USERS_SHEET);
  if (!sheet) return false;

  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim().toUpperCase() === userId.toUpperCase()) {
      // Mettre a jour la colonne G (titre)
      sheet.getRange(i + 1, 7).setValue(titre);
      Logger.log('Profil mis a jour: ' + userId + ' titre=' + titre);
      return true;
    }
  }
  return false;
}

/**
 * Cree l'onglet Users avec les utilisateurs par defaut
 */
function creerOngletUsers(ss) {
  var sheet = ss.insertSheet(CONFIG.USERS_SHEET);
  sheet.appendRow(['ID', 'Nom', 'Prenom', 'Telephone', 'Email', 'Role', 'Titre', 'SignatureId']);
  sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');

  // Utilisateurs par defaut (sans telephone pour eviter #ERROR!)
  sheet.appendRow(['ARNAUD', 'GUEDOU', 'Arnaud', '', '', 'vendeur', 'Responsable canton de Vaud', '']);
  sheet.appendRow(['LILOU', 'WOTQUENNE', 'Lilou', '', '', 'coordinateur', 'Coordinatrice', '']);

  // Ecrire les telephones en forcant le format texte cellule par cellule
  sheet.getRange('D2').setNumberFormat('@').setValue('+41 79 688 27 35');
  sheet.getRange('D3').setNumberFormat('@').setValue('022 827 36 97');

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 160);
  sheet.setColumnWidth(5, 200);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 220);
  sheet.setColumnWidth(8, 200);

  // Figer la premiere ligne
  sheet.setFrozenRows(1);

  return sheet;
}

// ============================================
// VENTILATION — Generation des lignes comptables
// ============================================

/**
 * Genere les lignes de ventilation a partir des postes d'un devis
 * Retourne un tableau de lignes : [{ n, ventil, poste, libelle, montant, dev }]
 */
/**
 * Parse une entrée ressource "2x Label [1/2 AM]" ou "2x Label"
 * Retourne { qty, nom, duree } ou null
 */
function parseRessourceEntry(item) {
  var match = String(item).match(/^(\d+)x\s+(.+?)(?:\s+\[(.+?)\])?$/);
  if (!match) return null;
  return { qty: parseInt(match[1]) || 1, nom: match[2].trim(), duree: match[3] || '' };
}

/**
 * Retourne le cout d'une ressource selon sa duree specifique
 * duree: '', '1/2 AM', '1/2 PM', '1j', '2j', etc.
 */
function getCoutAvecDuree(tarifs, type, nom, duree) {
  var items = (tarifs.items && tarifs.items[type]) || [];
  var found = null;
  for (var i = 0; i < items.length; i++) {
    if (items[i].item === nom) { found = items[i]; break; }
  }
  if (!found) return { cout: 0, jours: 1 };

  var coutJour = parseFloat(found.cout) || 0;
  var coutDemi = (found.coutDemiJour !== undefined && found.coutDemiJour !== null && found.coutDemiJour !== '')
    ? parseFloat(found.coutDemiJour) : coutJour / 2;

  var isDemi = (duree === '1/2 AM' || duree === '1/2 PM');
  if (isDemi) return { cout: coutDemi, jours: 0.5, isDemi: true };

  var matchJ = duree.match(/^(\d+)j$/);
  var jours = matchJ ? parseInt(matchJ[1]) : 0; // 0 = duree mission
  return { cout: coutJour, jours: jours, isDemi: false };
}

function genererVentilation(data) {
  var lignes = [];
  var tarifs = getTarifs();
  var n = 1;
  var V = CONFIG.VENTILATION;

  var postes = safeArray(data.postes).filter(function(p) { return p.mode === 'detail'; });

  postes.forEach(function(p) {
    var nbJoursVal = parseFloat(p.jours) || 1;
    var rdvs = safeArray(p.rdvs).filter(function(r) { return r.date; });
    var dateDebut = rdvs.length > 0 ? new Date(rdvs[0].date) : new Date();
    var joursOuvres = getJoursOuvres(dateDebut, nbJoursVal);
    var titre = safe(p.titre, 'PRESTATION').toUpperCase();

    // --- PERSONNEL : selon duree specifique ou mission ---
    var personnelItems = safeArray(p.personnel);
    personnelItems.forEach(function(item) {
      var parsed = parseRessourceEntry(item);
      if (!parsed) return;
      var qty = parsed.qty, nom = parsed.nom, duree = parsed.duree;
      var info = getCoutAvecDuree(tarifs, 'personnel', nom, duree);

      var abrNom = nom.replace(/Manutentionnaire du lourd/gi, 'HL')
                      .replace(/Manutentionnaire/gi, 'H')
                      .replace(/Chauffeur-Livreur/gi, 'CHAUF-LIV')
                      .replace(/Chauffeur/gi, 'CHAUF')
                      .replace(/Chef d'equipe/gi, 'CE');

      if (info.isDemi) {
        // Demi-journée spécifique : 1 seule ligne
        var dateFmt = joursOuvres.length > 0 ? Utilities.formatDate(joursOuvres[0], CONFIG.TIMEZONE, 'dd.MM.yyyy') : '';
        lignes.push({
          n: n++,
          ventil: V.personnel_demi.ventil,
          poste: V.personnel_demi.poste,
          libelle: dateFmt + ' - ' + qty + ' ' + abrNom + ' - ' + duree + ' - ' + titre,
          montant: info.cout * qty,
          dev: 'CHF'
        });
      } else if (info.jours > 0) {
        // Durée spécifique (ex: 2j) : lignes par jour ouvré sur cette durée
        var joursSpec = getJoursOuvres(dateDebut, info.jours);
        joursSpec.forEach(function(jourDate) {
          var df = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          lignes.push({
            n: n++,
            ventil: V.personnel.ventil,
            poste: V.personnel.poste,
            libelle: df + ' - ' + qty + ' ' + abrNom + ' - 1J - ' + titre,
            montant: info.cout * qty,
            dev: 'CHF'
          });
        });
      } else {
        // Durée mission (défaut) : 1 ligne par jour ouvré de la mission
        var hasDemiMission = (nbJoursVal % 1 !== 0);
        joursOuvres.forEach(function(jourDate, idx) {
          var df = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          var isLast = (idx === joursOuvres.length - 1);
          var isDemiLast = (isLast && hasDemiMission);
          var posteCode = isDemiLast ? V.personnel_demi : V.personnel;
          var coutLigne = isDemiLast ? (getCoutAvecDuree(tarifs, 'personnel', nom, '1/2 AM').cout * qty) : (info.cout * qty);
          var dureeLabel = isDemiLast ? '1/2J' : '1J';

          lignes.push({
            n: n++,
            ventil: posteCode.ventil,
            poste: posteCode.poste,
            libelle: df + ' - ' + qty + ' ' + abrNom + ' - ' + dureeLabel + ' - ' + titre,
            montant: coutLigne,
            dev: 'CHF'
          });
        });
      }
    });

    // --- VEHICULES : selon duree specifique ou mission ---
    var vehiculesItems = safeArray(p.vehicules);
    vehiculesItems.forEach(function(item) {
      var parsed = parseRessourceEntry(item);
      if (!parsed) return;
      var qty = parsed.qty, nom = parsed.nom, duree = parsed.duree;
      var info = getCoutAvecDuree(tarifs, 'vehicules', nom, duree);

      if (info.isDemi) {
        var dateFmt = joursOuvres.length > 0 ? Utilities.formatDate(joursOuvres[0], CONFIG.TIMEZONE, 'dd.MM.yyyy') : '';
        lignes.push({
          n: n++,
          ventil: V.vehicules.ventil,
          poste: V.vehicules.poste,
          libelle: dateFmt + ' - ' + qty + ' ' + nom + ' - ' + duree + ' - ' + titre,
          montant: info.cout * qty,
          dev: 'CHF'
        });
      } else if (info.jours > 0) {
        var joursSpec = getJoursOuvres(dateDebut, info.jours);
        joursSpec.forEach(function(jourDate) {
          var df = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          lignes.push({
            n: n++,
            ventil: V.vehicules.ventil,
            poste: V.vehicules.poste,
            libelle: df + ' - ' + qty + ' ' + nom + ' - 1J - ' + titre,
            montant: info.cout * qty,
            dev: 'CHF'
          });
        });
      } else {
        var hasDemiMission = (nbJoursVal % 1 !== 0);
        joursOuvres.forEach(function(jourDate, idx) {
          var df = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          var isLast = (idx === joursOuvres.length - 1);
          var isDemiLast = (isLast && hasDemiMission);
          var coutLigne = isDemiLast ? (getCoutAvecDuree(tarifs, 'vehicules', nom, '1/2 AM').cout * qty) : (info.cout * qty);
          var dureeLabel = isDemiLast ? '1/2J' : '1J';

          lignes.push({
            n: n++,
            ventil: V.vehicules.ventil,
            poste: V.vehicules.poste,
            libelle: df + ' - ' + qty + ' ' + nom + ' - ' + dureeLabel + ' - ' + titre,
            montant: coutLigne,
            dev: 'CHF'
          });
        });
      }
    });

    // --- ENGINS : selon duree specifique ou mission ---
    var enginsItems = safeArray(p.engins);
    enginsItems.forEach(function(item) {
      var match = String(item).match(/^(\d+)x\s+(.+?)(?:\s+\(\d+T\))?(?:\s+\[(.+?)\])?$/);
      if (!match) return;
      var qty = parseInt(match[1]) || 1;
      var nom = match[2].trim();
      var duree = match[3] || '';
      var info = getCoutAvecDuree(tarifs, 'engins', nom, duree);
      var enginConf = V.engins[nom] || V.engins._default;

      if (info.isDemi) {
        var dateFmt = joursOuvres.length > 0 ? Utilities.formatDate(joursOuvres[0], CONFIG.TIMEZONE, 'dd.MM.yyyy') : '';
        lignes.push({
          n: n++,
          ventil: enginConf.ventil,
          poste: enginConf.poste,
          libelle: dateFmt + ' - ' + qty + ' ' + nom + ' - ' + duree + ' - ' + titre,
          montant: info.cout * qty,
          dev: 'CHF'
        });
      } else if (info.jours > 0) {
        var joursSpec = getJoursOuvres(dateDebut, info.jours);
        joursSpec.forEach(function(jourDate) {
          var df = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          lignes.push({
            n: n++,
            ventil: enginConf.ventil,
            poste: enginConf.poste,
            libelle: df + ' - ' + qty + ' ' + nom + ' - 1J - ' + titre,
            montant: info.cout * qty,
            dev: 'CHF'
          });
        });
      } else {
        var hasDemiMission = (nbJoursVal % 1 !== 0);
        joursOuvres.forEach(function(jourDate, idx) {
          var df = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          var isLast = (idx === joursOuvres.length - 1);
          var isDemiLast = (isLast && hasDemiMission);
          var coutLigne = isDemiLast ? (getCoutAvecDuree(tarifs, 'engins', nom, '1/2 AM').cout * qty) : (info.cout * qty);
          var dureeLabel = isDemiLast ? '1/2J' : '1J';

          lignes.push({
            n: n++,
            ventil: enginConf.ventil,
            poste: enginConf.poste,
            libelle: df + ' - ' + qty + ' ' + nom + ' - ' + dureeLabel + ' - ' + titre,
            montant: coutLigne,
            dev: 'CHF'
          });
        });
      }
    });

    // --- MATERIEL : 1 seule ligne (pieces) ---
    var materielItems = safeArray(p.materiel);
    materielItems.forEach(function(item) {
      var parsed = parseRessourceEntry(item);
      if (!parsed) return;
      var qty = parsed.qty, nom = parsed.nom;

      var coutUnit = 0;
      var unite = 'piece';
      if (tarifs.items.materiel) {
        var found = tarifs.items.materiel.find(function(t) { return t.item === nom; });
        if (found) { coutUnit = parseFloat(found.cout) || 0; unite = found.unite || 'piece'; }
      }

      var montant = (unite === 'jour') ? coutUnit * qty * nbJoursVal : coutUnit * qty;

      lignes.push({
        n: n++,
        ventil: V.materiel_emballage.ventil,
        poste: V.materiel_emballage.poste,
        libelle: qty + 'x ' + nom + ' - ' + titre,
        montant: montant,
        dev: 'CHF'
      });
    });
  });

  // Total charges locales
  var totalLocal = 0;
  lignes.forEach(function(l) { totalLocal += l.montant; });
  lignes.push({
    n: n++,
    ventil: 'FR-LOC-TOT',
    poste: '',
    libelle: 'TOTAL DES CHARGES LOCALES',
    montant: totalLocal,
    dev: 'CHF',
    isTotal: true
  });

  return lignes;
}

// ============================================
// TARIFS & MARGES (onglet Tarifs)
// ============================================
var TARIFS_SHEET_NAME = 'Tarifs';

/**
 * Retourne tous les tarifs et marges
 * Format retour : { items: { personnel: [{item, cout, unite},...], ... }, marges: { personnel: 35, ... } }
 */
function getTarifs() {
  var ss = getSs();
  if (!ss) return { items: {}, marges: {} };
  var sheet = ss.getSheetByName(TARIFS_SHEET_NAME);
  if (!sheet) {
    // Creer l'onglet avec des valeurs par defaut
    sheet = creerOngletTarifs(ss);
  }

  var values = sheet.getDataRange().getValues();
  var items = {};
  var marges = {};

  // En-tetes : Type | Item | Cout unitaire | Unite | Marge % | Cout demi-jour
  for (var i = 1; i < values.length; i++) {
    var type = String(values[i][0] || '').trim();
    var item = String(values[i][1] || '').trim();
    var cout = parseFloat(values[i][2]) || 0;
    var unite = String(values[i][3] || 'jour').trim();
    var marge = parseFloat(values[i][4]);
    var coutDemiJour = (values[i].length > 5 && values[i][5] !== '' && values[i][5] !== null) ? parseFloat(values[i][5]) : null;
    var coutKm = (values[i].length > 6 && values[i][6] !== '' && values[i][6] !== null) ? parseFloat(values[i][6]) : null;

    if (!type) continue;

    if (type === '_marge') {
      // Ligne de marge par categorie
      marges[item] = cout;
    } else {
      // Ligne de tarif
      if (!items[type]) items[type] = [];
      var entry = { item: item, cout: cout, unite: unite };
      if (coutDemiJour !== null && !isNaN(coutDemiJour)) entry.coutDemiJour = coutDemiJour;
      if (coutKm !== null && !isNaN(coutKm)) entry.coutKm = coutKm;
      items[type].push(entry);
      // Si une marge est definie sur la ligne, l'utiliser comme defaut categorie
      if (!isNaN(marge) && !marges[type]) marges[type] = marge;
    }
  }

  return { items: items, marges: marges };
}

/**
 * Sauvegarde les tarifs (ecrase l'onglet)
 * data = { items: { personnel: [{item, cout, unite},...], ... }, marges: { personnel: 35, ... } }
 */
function saveTarifs(data) {
  var ss = getSs();
  if (!ss) return false;
  var sheet = ss.getSheetByName(TARIFS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(TARIFS_SHEET_NAME);

  // Effacer le contenu existant
  sheet.clearContents();

  // En-tetes
  sheet.appendRow(['Type', 'Item', 'Cout unitaire', 'Unite', 'Marge %', 'Cout 1/2 jour', 'CHF/km']);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');

  // Marges par categorie
  var marges = data.marges || {};
  for (var cat in marges) {
    sheet.appendRow(['_marge', cat, marges[cat], '%', '', '', '']);
  }

  // Items
  var items = data.items || {};
  for (var type in items) {
    var list = items[type] || [];
    for (var i = 0; i < list.length; i++) {
      var demiJour = (list[i].coutDemiJour !== undefined && list[i].coutDemiJour !== null) ? list[i].coutDemiJour : '';
      var km = (list[i].coutKm !== undefined && list[i].coutKm !== null) ? list[i].coutKm : '';
      sheet.appendRow([type, list[i].item, list[i].cout, list[i].unite || 'jour', marges[type] || '', demiJour, km]);
    }
  }

  // Formater les colonnes
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 120);

  return true;
}

/**
 * Cree l'onglet Tarifs avec des valeurs par defaut
 */
function creerOngletTarifs(ss) {
  var sheet = ss.insertSheet(TARIFS_SHEET_NAME);
  sheet.appendRow(['Type', 'Item', 'Cout unitaire', 'Unite', 'Marge %', 'Cout 1/2 jour', 'CHF/km']);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');

  // Marges par defaut
  sheet.appendRow(['_marge', 'personnel', 35, '%', '', '', '']);
  sheet.appendRow(['_marge', 'vehicules', 25, '%', '', '', '']);
  sheet.appendRow(['_marge', 'engins', 40, '%', '', '', '']);
  sheet.appendRow(['_marge', 'materiel', 10, '%', '', '', '']);

  // Tarifs par defaut — cout journée, cout demi-journée et CHF/km (source: grille Pelichet)
  // Format: [type, item, cout/jour, unite, '', cout/demi-jour, CHF/km]
  var defauts = [
    ['personnel', 'Manutentionnaire', 380, 'jour', '', 210, ''],
    ['personnel', 'Manutentionnaire du lourd', 400, 'jour', '', 220, ''],
    ['personnel', 'Chauffeur', 380, 'jour', '', 210, ''],
    ['personnel', "Chef d'equipe", 450, 'jour', '', 250, ''],
    ['personnel', 'CHAUF-LIVREUR', 400, 'jour', '', 220, ''],
    ['personnel', 'emballeur', 380, 'jour', '', 210, ''],
    ['vehicules', 'VL', 150, 'jour', '', 75, 0.8],
    ['vehicules', '1F', 150, 'jour', '', 75, 0.8],
    ['vehicules', 'PL', 350, 'jour', '', 250, 1.5],
    ['vehicules', 'Semi', 500, 'jour', '', 500, 1.7],
    ['vehicules', 'Box IT', 150, 'jour', '', 75, 0.8],
    ['vehicules', 'PL avec hayon', 350, 'jour', '', 250, 1.5],
    ['engins', 'Chariot elevateur', 600, 'jour', '', 400, ''],
    ['engins', 'Grue mobile', 1200, 'jour', '', 800, ''],
    ['engins', 'Monte-meuble', 400, 'jour', '', 300, ''],
    ['materiel', 'Chariots', 5, 'piece', '', '', ''],
    ['materiel', 'Rouleaux bulle', 15, 'piece', '', '', ''],
    ['materiel', 'Adhesif', 3, 'piece', '', '', ''],
    ['materiel', 'Transpalette', 30, 'jour', '', 20, '']
  ];
  defauts.forEach(function(row) { sheet.appendRow(row); });

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 90);

  return sheet;
}

// ============================================
// VALIDATION SERVEUR
// ============================================
function validateData(data) {
  const errors = [];
  if (!data.ref || !String(data.ref).trim()) errors.push('Référence manquante');
  if (!data.client || !String(data.client).trim()) errors.push('Client manquant');
  if (!data.montantHT || isNaN(parseFloat(data.montantHT)) || parseFloat(data.montantHT) <= 0) {
    errors.push('Montant HT invalide');
  }
  if (!data.adresseDepart) errors.push('Adresse de départ manquante');
  if (!data.adresseArrivee) errors.push('Adresse d\'arrivée manquante');
  return errors;
}

// ============================================
// RECHERCHE DOSSIER (GET)
// ============================================
function doGet(e) {
  try {
    const action = (e.parameter && e.parameter.action) || '';

    if (action === 'users_get') {
      var users = getUsers();
      return jsonResponse({ status: 'success', data: users });
    }

    if (action === 'user_add') {
      try {
        var newUserData = JSON.parse(e.parameter.data || '{}');
        var newUser = addUser(newUserData);
        if (newUser) {
          return jsonResponse({ status: 'success', data: newUser });
        }
        return jsonResponse({ status: 'error', message: 'Nom et prenom requis ou doublon' });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Erreur creation: ' + err });
      }
    }

    if (action === 'user_update_profile') {
      try {
        var profData = JSON.parse(e.parameter.data || '{}');
        var profUserId = profData.userId || '';
        var profTitre = profData.titre || '';
        if (!profUserId) return jsonResponse({ status: 'error', message: 'userId requis' });
        var updated = updateUserProfile(profUserId, profTitre);
        return jsonResponse({ status: updated ? 'success' : 'error' });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Erreur profil: ' + err });
      }
    }

    if (action === 'user_upload_signature') {
      try {
        var sigUserId = e.parameter.userId || '';
        var sigData = e.parameter.data || '';
        if (!sigUserId || !sigData) return jsonResponse({ status: 'error', message: 'userId et data requis' });
        var sigId = updateUserSignature(sigUserId, sigData);
        if (sigId) {
          return jsonResponse({ status: 'success', signatureId: sigId });
        }
        return jsonResponse({ status: 'error', message: 'Utilisateur introuvable' });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Erreur upload signature: ' + err });
      }
    }

    if (action === 'rechercher') {
      const ref = e.parameter.ref || '';
      const userId = e.parameter.user || '';
      const result = rechercherDossier(ref, userId);
      if (result) {
        return jsonResponse({ status: 'success', data: result });
      }
      return jsonResponse({ status: 'not_found' });
    }

    if (action === 'lister') {
      const userId = e.parameter.user || '';
      const dossiers = listerDossiers(userId);
      return jsonResponse({ status: 'success', data: dossiers });
    }

    if (action === 'custom_get') {
      var items = getCustomItems();
      return jsonResponse({ status: 'success', data: items });
    }

    if (action === 'custom_add') {
      var type = e.parameter.type || '';
      var value = e.parameter.value || '';
      if (!type || !value) return jsonResponse({ status: 'error', message: 'type et value requis' });
      var added = addCustomItem(type, value);
      return jsonResponse({ status: 'success', added: added });
    }

    if (action === 'tarifs_get') {
      var tarifs = getTarifs();
      return jsonResponse({ status: 'success', data: tarifs });
    }

    if (action === 'tarifs_save') {
      try {
        var tarifsData = JSON.parse(e.parameter.data || '{}');
        var saved = saveTarifs(tarifsData);
        return jsonResponse({ status: saved ? 'success' : 'error' });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'JSON invalide: ' + err });
      }
    }

    if (action === 'km_calc') {
      try {
        var depart = e.parameter.depart || '';
        var arrivee = e.parameter.arrivee || '';
        if (!depart || !arrivee) return jsonResponse({ status: 'error', message: 'Adresses depart et arrivee requises' });
        var result = calculerKilometrage(depart, arrivee);
        if (result.error) return jsonResponse({ status: 'error', message: result.error });
        return jsonResponse({ status: 'success', data: result });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Calcul KM: ' + err });
      }
    }

    if (action === 'company_search') {
      var query = e.parameter.q || '';
      if (query.length < 2) return jsonResponse({ status: 'success', results: [] });
      try {
        var results = searchSwissCompany(query);
        return jsonResponse({ status: 'success', results: results });
      } catch (err) {
        return jsonResponse({ status: 'error', message: err.toString(), results: [] });
      }
    }

    if (action === 'ics') {
      try {
        var icsUserId = e.parameter.user || '';
        if (!icsUserId) return ContentService.createTextOutput('userId requis').setMimeType(ContentService.MimeType.TEXT);
        var icsContent = generateICS(icsUserId);
        return ContentService.createTextOutput(icsContent).setMimeType(ContentService.MimeType.TEXT);
      } catch (err) {
        return ContentService.createTextOutput('Erreur: ' + err).setMimeType(ContentService.MimeType.TEXT);
      }
    }

    if (action === 'calendar_data') {
      try {
        var calUserId = e.parameter.user || '';
        if (!calUserId) return jsonResponse({ status: 'error', message: 'userId requis' });
        var calData = getCalendarData(calUserId);
        return jsonResponse({ status: 'success', data: calData });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Calendar: ' + err });
      }
    }

    if (action === 'planning_generate') {
      try {
        var planUserId = e.parameter.user || '';
        if (!planUserId) return jsonResponse({ status: 'error', message: 'userId requis' });
        var planResult = genererPlanning(planUserId);
        return jsonResponse(planResult);
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Erreur planning: ' + err });
      }
    }

    if (action === 'ventilation') {
      try {
        var ventilData = JSON.parse(e.parameter.data || '{}');
        var lignes = genererVentilation(ventilData);
        return jsonResponse({ status: 'success', data: lignes });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Ventilation: ' + err });
      }
    }

    return jsonResponse({ status: 'error', message: 'Action inconnue' });
  } catch (err) {
    Logger.log('ERREUR doGet: ' + err);
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Trouve la feuille de suivi pour un utilisateur (par ID ou globale)
 * Cherche d'abord Suivi_Devis_{userId}, puis Suivi_Devis, puis la premiere feuille
 */
function getSheet(userId) {
  const ss = getSs();
  if (!ss) return null;

  // 1. Onglet utilisateur specifique
  if (userId) {
    var userSheet = ss.getSheetByName(getUserSheetName(userId));
    if (userSheet) return userSheet;
  }

  // 2. Onglet generique
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (sheet) return sheet;

  // 3. Fallback premiere feuille
  const sheets = ss.getSheets();
  if (sheets.length === 0) return null;
  sheet = sheets[0];
  Logger.log('Feuille "' + CONFIG.SHEET_NAME + '" introuvable, utilisation de "' + sheet.getName() + '"');
  return sheet;
}

/**
 * Construit un index des colonnes à partir des en-têtes.
 * Matching par mots-clés contenus dans l'en-tête (insensible casse/accents).
 * Retourne { ref: 1, client: 5, ... }
 */
function buildColumnIndex(headers) {
  const map = {};

  // Règles : si l'en-tête contient un de ces mots-clés -> associer à cette clé
  // L'ordre compte : la première règle qui matche gagne
  var rules = [
    { key: 'date',    words: ['date'] },
    { key: 'ref',     words: ['ref', 'réf'] },
    { key: 'societe', words: ['societ', 'sociét'] },
    { key: 'email',   words: ['email', 'e-mail', 'mail'] },
    { key: 'statut',  words: ['statut'] },
    { key: 'client',  words: ['client'] },
    { key: 'facturation', words: ['facturation', 'adresse fact'] },
    { key: 'contact', words: ['contact'] },
    { key: 'depart',  words: ['depart', 'départ', 'enlev', 'enlèv'] },
    { key: 'arrivee', words: ['arriv'] },
    { key: 'prevu',   words: ['prevu', 'prévu'] },
    { key: 'detailTarife', words: ['detail', 'détail', 'tarif'] },
    { key: 'ht',      words: ['ht'] },
    { key: 'ttc',     words: ['ttc'] },
    { key: 'lien',    words: ['lien', 'url', 'pdf'] },
    { key: 'json',    words: ['json', 'data'] }
  ];

  headers.forEach(function(h, idx) {
    var label = String(h).trim().toLowerCase();
    for (var r = 0; r < rules.length; r++) {
      if (map[rules[r].key] !== undefined) continue; // Déjà trouvé
      for (var w = 0; w < rules[r].words.length; w++) {
        if (label.indexOf(rules[r].words[w]) !== -1) {
          map[rules[r].key] = idx;
          break;
        }
      }
    }
  });

  Logger.log('Column index: ' + JSON.stringify(map));
  return map;
}

/**
 * Extrait une valeur d'une ligne par nom de colonne
 */
function col(row, colMap, name, fallback) {
  if (colMap[name] === undefined) return fallback || '';
  return safe(row[colMap[name]], fallback);
}

/**
 * Liste tous les dossiers (ref, client, date, montant, statut)
 * Retourne les 200 derniers, du plus récent au plus ancien
 */
function listerDossiers(userId) {
  const sheet = getSheet(userId);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const colMap = buildColumnIndex(values[0]);
  const refIdx = colMap.ref !== undefined ? colMap.ref : 1;

  // Dédupliquer par référence : on parcourt du plus récent au plus ancien,
  // seule la première occurrence (= la plus récente) est gardée
  const seen = {};
  const dossiers = [];

  for (let i = values.length - 1; i >= 1 && dossiers.length < 200; i--) {
    const ref = String(values[i][refIdx] || '').trim();
    if (!ref) continue;
    if (seen[ref]) continue; // Déjà vu → doublon, on saute
    seen[ref] = true;

    const dateVal = colMap.date !== undefined ? values[i][colMap.date] : '';
    let dateFmt = '';
    if (dateVal instanceof Date) {
      dateFmt = Utilities.formatDate(dateVal, CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
    } else {
      dateFmt = String(dateVal);
    }

    dossiers.push({
      ref: ref,
      client: col(values[i], colMap, 'client'),
      date: dateFmt,
      statut: col(values[i], colMap, 'statut'),
      montantHT: parseFloat(col(values[i], colMap, 'ht', '0')) || 0
    });
  }

  return dossiers;
}

function rechercherDossier(ref, userId) {
  const sheet = getSheet(userId);
  if (!sheet) return null;

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return null;

  const colMap = buildColumnIndex(values[0]);
  const refIdx = colMap.ref !== undefined ? colMap.ref : 1;

  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][refIdx]).trim().toUpperCase() === ref.trim().toUpperCase()) {

      // Chercher du JSON dans toutes les colonnes (le format d'archivage peut varier)
      // Priorité : colonne "json" mappée > scan de toutes les colonnes
      var jsonFound = null;
      var searchOrder = [];
      if (colMap.json !== undefined) searchOrder.push(colMap.json);
      // Scan depuis la dernière colonne (plus probable)
      for (var c = values[i].length - 1; c >= 0; c--) {
        if (c !== colMap.json) searchOrder.push(c);
      }
      for (var s = 0; s < searchOrder.length; s++) {
        var cellVal = values[i][searchOrder[s]];
        if (cellVal && String(cellVal).trim().startsWith('{"')) {
          try {
            jsonFound = JSON.parse(cellVal);
            Logger.log('JSON trouvé en colonne ' + searchOrder[s] + ' pour ref ' + ref);
            break;
          } catch (e) { /* pas du JSON valide, continuer */ }
        }
      }

      if (jsonFound) return jsonFound;

      // Reconstruction dynamique depuis les en-têtes détectés
      return {
        ref: col(values[i], colMap, 'ref'),
        typeSociete: col(values[i], colMap, 'societe', 'Pelichet'),
        client: col(values[i], colMap, 'client'),
        adresseClient: col(values[i], colMap, 'facturation'),
        contact: col(values[i], colMap, 'contact'),
        adresseDepart: col(values[i], colMap, 'depart'),
        adresseArrivee: col(values[i], colMap, 'arrivee'),
        datePrevue: col(values[i], colMap, 'prevu'),
        montantHT: parseFloat(col(values[i], colMap, 'ht', '0')) || 0,
        genre: 'Monsieur',
        postes: []
      };
    }
  }
  return null;
}

// ============================================
// PLANNING CONSOLIDÉ
// ============================================

/**
 * Jours fériés suisses (Genève) — version Set de strings 'YYYY-MM-DD'
 * Utilisée par le planning consolidé (ne pas confondre avec getJoursFeries qui retourne des Date)
 */
function getPlanningFeriesSet(year) {
  var feries = [];
  // Fériés fixes
  feries.push(year + '-01-01'); // Nouvel An
  feries.push(year + '-01-02'); // 2 janvier
  feries.push(year + '-05-01'); // Fête du Travail
  feries.push(year + '-08-01'); // Fête nationale
  feries.push(year + '-12-25'); // Noël
  feries.push(year + '-12-31'); // Restauration de la République (GE)

  // Pâques (algorithme de Gauss)
  var a = year % 19;
  var b = Math.floor(year / 100);
  var c = year % 100;
  var d = Math.floor(b / 4);
  var e2 = b % 4;
  var f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4);
  var k = c % 4;
  var l = (32 + 2 * e2 + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var month = Math.floor((h + l - 7 * m + 114) / 31);
  var day = ((h + l - 7 * m + 114) % 31) + 1;
  var easter = new Date(year, month - 1, day);

  function addDays(d, n) { var r = new Date(d); r.setDate(r.getDate() + n); return r; }
  function fmt(d) {
    var mm = String(d.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(d.getDate()); if (dd.length < 2) dd = '0' + dd;
    return d.getFullYear() + '-' + mm + '-' + dd;
  }

  feries.push(fmt(addDays(easter, -2)));  // Vendredi Saint
  feries.push(fmt(addDays(easter, 1)));   // Lundi de Pâques
  feries.push(fmt(addDays(easter, 39)));  // Ascension
  feries.push(fmt(addDays(easter, 50)));  // Lundi de Pentecôte

  // Jeûne genevois
  var sept1 = new Date(year, 8, 1);
  while (sept1.getDay() !== 0) sept1.setDate(sept1.getDate() + 1);
  sept1.setDate(sept1.getDate() + 4);
  feries.push(fmt(sept1));

  var set = {};
  feries.forEach(function(f) { set[f] = true; });
  return set;
}

/**
 * Nom du jour férié (pour affichage)
 */
function getNomFerie(dateStr, year) {
  var feries = {};
  feries[year + '-01-01'] = 'Nouvel An';
  feries[year + '-01-02'] = '2 Janvier';
  feries[year + '-03-01'] = 'Instauration République';
  feries[year + '-05-01'] = 'Fête du Travail';
  feries[year + '-08-01'] = 'Fête nationale';
  feries[year + '-12-25'] = 'Noël';
  feries[year + '-12-31'] = 'Restauration République';

  // Pâques
  var a = year % 19;
  var b = Math.floor(year / 100);
  var c = year % 100;
  var d = Math.floor(b / 4);
  var e2 = b % 4;
  var f = Math.floor((b + 8) / 25);
  var g = Math.floor((b - f + 1) / 3);
  var h = (19 * a + b - d - g + 15) % 30;
  var i = Math.floor(c / 4);
  var k = c % 4;
  var l = (32 + 2 * e2 + 2 * i - h - k) % 7;
  var m = Math.floor((a + 11 * h + 22 * l) / 451);
  var month = Math.floor((h + l - 7 * m + 114) / 31);
  var day = ((h + l - 7 * m + 114) % 31) + 1;
  var easter = new Date(year, month - 1, day);
  function addDays(d, n) { var r = new Date(d); r.setDate(r.getDate() + n); return r; }
  function fmt(d) {
    var mm = String(d.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(d.getDate()); if (dd.length < 2) dd = '0' + dd;
    return d.getFullYear() + '-' + mm + '-' + dd;
  }
  feries[fmt(addDays(easter, -2))] = 'Vendredi Saint';
  feries[fmt(addDays(easter, 1))] = 'Lundi de Pâques';
  feries[fmt(addDays(easter, 39))] = 'Ascension';
  feries[fmt(addDays(easter, 50))] = 'Lundi de Pentecôte';

  return feries[dateStr] || '';
}

/**
 * Génère les jours ouvrés à partir d'une date de début + nb jours
 */
function genererJoursOuvres(dateDebut, nbJours, feriesSet) {
  var jours = [];
  var current = new Date(dateDebut);
  var count = 0;
  var maxIter = nbJours * 3 + 30; // sécurité

  while (count < nbJours && maxIter > 0) {
    maxIter--;
    var dow = current.getDay(); // 0=dim, 6=sam
    var mm = String(current.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(current.getDate()); if (dd.length < 2) dd = '0' + dd;
    var key = current.getFullYear() + '-' + mm + '-' + dd;

    if (dow !== 0 && dow !== 6 && !feriesSet[key]) {
      jours.push(new Date(current));
      count++;
    }
    current.setDate(current.getDate() + 1);
  }
  return jours;
}

/**
 * Retourne les données de calendrier pour un utilisateur (vue agrégée par jour)
 * Format retour : { entries: [{date, ref, client, titre, jours, effectif, personnel, vehicules, engins, materiel, tache}], minDate, maxDate }
 */
function getCalendarData(userId) {
  var sheet = getSheet(userId);
  if (!sheet) return { entries: [], minDate: null, maxDate: null };
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { entries: [], minDate: null, maxDate: null };

  // Déduplication par ref (garder la version la plus récente)
  var colMap = buildColumnIndex(values[0]);
  var refIdx = colMap.ref !== undefined ? colMap.ref : 1;
  var seenRefs = {};
  var latestRows = [];
  for (var i = values.length - 1; i >= 1; i--) {
    var ref = String(values[i][refIdx] || '').trim();
    if (!ref || seenRefs[ref]) continue;
    seenRefs[ref] = true;
    latestRows.push(values[i]);
  }

  // Collecter toutes les années pour les jours fériés
  var yearsSet = {};
  var allEntries = [];

  latestRows.forEach(function(row) {
    // Chercher le JSON dans les colonnes
    var data = null;
    for (var c = row.length - 1; c >= 0; c--) {
      var cellVal = row[c];
      if (cellVal && String(cellVal).trim().startsWith('{"')) {
        try { data = JSON.parse(cellVal); break; } catch (e) { /* skip */ }
      }
    }
    if (!data) return;

    var postes = safeArray(data.postes).filter(function(p) { return p.mode === 'detail'; });
    var client = safe(data.client, 'Client');
    var ref = safe(data.ref, 'SANS_REF');

    postes.forEach(function(p) {
      var rdvs = safeArray(p.rdvs).filter(function(r) { return r.date; });
      if (rdvs.length === 0) return;
      var rawDate = rdvs[0].date;
      var dateDebut = parseDateSafe(rawDate);
      if (!dateDebut) return;
      yearsSet[dateDebut.getFullYear()] = true;

      var nbJoursVal = parseFloat(p.jours) || 1;
      var joursOuvres = getJoursOuvres(dateDebut, Math.ceil(nbJoursVal));

      // Extraction des ressources
      var personnelList = safeArray(p.personnel);
      var vehiculesList = safeArray(p.vehicules);
      var enginsList = safeArray(p.engins);
      var materielList = safeArray(p.materiel);

      var effectif = 0;
      personnelList.forEach(function(s) {
        var m = String(s).match(/^(\d+)x/);
        if (m) effectif += parseInt(m[1]);
      });

      joursOuvres.forEach(function(jour) {
        var dateStr = Utilities.formatDate(jour, CONFIG.TIMEZONE, 'yyyy-MM-dd');
        allEntries.push({
          date: dateStr,
          ref: ref,
          client: client,
          titre: safe(p.titre, 'Prestation'),
          jours: nbJoursVal,
          effectif: effectif,
          personnel: transformerPersonnel(personnelList),
          vehicules: vehiculesList,
          engins: enginsList,
          materiel: materielList,
          tache: safe(p.tache)
        });
      });
    });
  });

  // Trier par date
  allEntries.sort(function(a, b) { return a.date < b.date ? -1 : 1; });

  var minDate = allEntries.length > 0 ? allEntries[0].date : null;
  var maxDate = allEntries.length > 0 ? allEntries[allEntries.length - 1].date : null;

  return { entries: allEntries, minDate: minDate, maxDate: maxDate };
}

/**
 * Exporte un Google Doc en DOCX (Word) via l'API Drive.
 * Retourne le Blob DOCX ou null en cas d'erreur.
 */
function exportDocAsDocx(docId) {
  try {
    // Laisser le temps à Drive d'indexer le document après saveAndClose
    Utilities.sleep(800);

    var url = 'https://docs.google.com/document/d/' + docId + '/export?format=docx';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
      followRedirects: true
    });
    var code = response.getResponseCode();
    Logger.log('exportDocAsDocx docId=' + docId + ' code=' + code);

    if (code === 200) {
      var blob = response.getBlob();
      blob.setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      blob.setName('export.docx');
      Logger.log('DOCX export OK, size=' + blob.getBytes().length);
      return blob;
    }

    // Tentative alternative : API Drive v3 export
    var altUrl = 'https://www.googleapis.com/drive/v3/files/' + docId
      + '/export?mimeType=' + encodeURIComponent('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    var altResp = UrlFetchApp.fetch(altUrl, {
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true,
      followRedirects: true
    });
    var altCode = altResp.getResponseCode();
    Logger.log('exportDocAsDocx fallback code=' + altCode);
    if (altCode === 200) {
      var b2 = altResp.getBlob();
      b2.setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
      return b2;
    }

    Logger.log('exportDocAsDocx body: ' + response.getContentText().substring(0, 300));
  } catch (e) {
    Logger.log('exportDocAsDocx error: ' + e);
  }
  return null;
}

/**
 * Retourne (ou crée) le sous-dossier Drive dédié à un utilisateur.
 * Structure : FOLDER_ID/<USERID>/
 * Si userId vide ou introuvable, retourne le dossier racine.
 */
function getUserFolder(userId) {
  var rootFolder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  if (!userId) return rootFolder;

  var folderName = String(userId).toUpperCase().trim();
  if (!folderName) return rootFolder;

  // Chercher le sous-dossier existant
  var existing = rootFolder.getFoldersByName(folderName);
  if (existing.hasNext()) return existing.next();

  // Créer le sous-dossier
  var newFolder = rootFolder.createFolder(folderName);
  Logger.log('Sous-dossier utilisateur créé : ' + folderName + ' (ID: ' + newFolder.getId() + ')');
  return newFolder;
}

/**
 * Génère un flux iCalendar (.ics) à partir des missions d'un utilisateur.
 * Compatible avec iPhone Calendar, Google Calendar, Outlook.
 * Chaque prestation devient un événement sur toute sa durée (en jours).
 */
function generateICS(userId) {
  var calData = getCalendarData(userId);
  var entries = calData.entries || [];

  // Regrouper les entrées consécutives d'une même prestation (même ref + titre)
  // en un seul événement multi-jours
  var groups = {};
  entries.forEach(function(e) {
    var key = e.ref + '|' + e.titre;
    if (!groups[key]) groups[key] = { ref: e.ref, titre: e.titre, client: e.client, dates: [], entry: e };
    groups[key].dates.push(e.date);
  });

  function esc(s) {
    return String(s || '')
      .replace(/\\/g, '\\\\')
      .replace(/;/g, '\\;')
      .replace(/,/g, '\\,')
      .replace(/\n/g, '\\n')
      .replace(/\r/g, '');
  }
  function fmtDate(s) {
    // 'YYYY-MM-DD' -> 'YYYYMMDD'
    return s.replace(/-/g, '');
  }
  function addDay(dateStr) {
    var parts = dateStr.split('-');
    var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]) + 1);
    var mm = String(d.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(d.getDate()); if (dd.length < 2) dd = '0' + dd;
    return d.getFullYear() + '-' + mm + '-' + dd;
  }

  var now = new Date();
  var dtstamp = Utilities.formatDate(now, 'UTC', "yyyyMMdd'T'HHmmss'Z'");

  var lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//Pelichet NLC//Devis & Logistique//FR',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'X-WR-CALNAME:Pelichet — ' + userId,
    'X-WR-TIMEZONE:Europe/Zurich',
    'X-PUBLISHED-TTL:PT6H',
    'REFRESH-INTERVAL;VALUE=DURATION:PT6H'
  ];

  Object.keys(groups).forEach(function(key) {
    var g = groups[key];
    g.dates.sort();
    var dateDebut = g.dates[0];
    var dateFin = g.dates[g.dates.length - 1];
    var dtend = addDay(dateFin); // ICS DTEND pour all-day = lendemain du dernier jour

    var e = g.entry;
    var descLines = [];
    descLines.push('Client: ' + g.client);
    descLines.push('Prestation: ' + g.titre);
    descLines.push('Durée: ' + (e.jours > 1 ? (e.jours + ' jours') : '1 jour'));
    if (e.personnel && e.personnel.length) descLines.push('Personnel: ' + e.personnel.join(', '));
    if (e.vehicules && e.vehicules.length) descLines.push('Véhicules: ' + e.vehicules.join(', '));
    if (e.engins && e.engins.length) descLines.push('Véhicules spéciaux: ' + e.engins.join(', '));
    if (e.materiel && e.materiel.length) descLines.push('Matériel: ' + e.materiel.join(', '));
    if (e.tache) descLines.push('\nInstructions:\n' + e.tache);

    var summary = g.client + ' — ' + g.titre;
    if (e.effectif) summary += ' (' + e.effectif + 'H)';

    var uid = (g.ref + '-' + fmtDate(dateDebut) + '@pelichet.ch').replace(/\s+/g, '');

    lines.push('BEGIN:VEVENT');
    lines.push('UID:' + uid);
    lines.push('DTSTAMP:' + dtstamp);
    lines.push('DTSTART;VALUE=DATE:' + fmtDate(dateDebut));
    lines.push('DTEND;VALUE=DATE:' + fmtDate(dtend));
    lines.push('SUMMARY:' + esc(summary));
    lines.push('DESCRIPTION:' + esc(descLines.join('\n')));
    lines.push('CATEGORIES:Pelichet,Logistique');
    lines.push('STATUS:CONFIRMED');
    lines.push('TRANSP:OPAQUE');
    lines.push('END:VEVENT');
  });

  lines.push('END:VCALENDAR');

  // ICS : lignes séparées par CRLF
  return lines.join('\r\n');
}

/**
 * Parse une date robustement (ISO, dd.MM.yyyy, etc.)
 */
function parseDateSafe(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v.getTime()) ? null : new Date(v.getFullYear(), v.getMonth(), v.getDate(), 12, 0, 0);
  var s = String(v).trim();
  var iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(parseInt(iso[1]), parseInt(iso[2]) - 1, parseInt(iso[3]), 12, 0, 0);
  var fr = s.match(/^(\d{1,2})[\.\/](\d{1,2})[\.\/](\d{4})/);
  if (fr) return new Date(parseInt(fr[3]), parseInt(fr[2]) - 1, parseInt(fr[1]), 12, 0, 0);
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Génère le planning consolidé Google Sheet pour un utilisateur.
 * Lit tous les dossiers, extrait les postes détail avec dates,
 * et crée un Sheet avec colonnes par client.
 */
function genererPlanning(userId) {
  var sheet = getSheet(userId);
  if (!sheet) return { status: 'error', message: 'Aucun dossier trouvé' };

  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return { status: 'error', message: 'Aucun dossier' };

  // Récupérer tous les dossiers avec JSON
  var dossiers = [];
  for (var i = 1; i < values.length; i++) {
    var jsonStr = null;
    for (var c = values[i].length - 1; c >= 0; c--) {
      var cellVal = values[i][c];
      if (cellVal && String(cellVal).trim().startsWith('{"')) {
        try {
          jsonStr = JSON.parse(cellVal);
          break;
        } catch (e) { /* pas du JSON */ }
      }
    }
    if (!jsonStr) continue;

    var postes = jsonStr.postes || [];
    var postesDetail = postes.filter(function(p) { return p.mode === 'detail'; });

    // Vérifier qu'au moins un poste a une date RDV
    var hasDate = postesDetail.some(function(p) {
      return p.rdvs && p.rdvs.length > 0 && p.rdvs[0].date;
    });
    if (!hasDate || postesDetail.length === 0) continue;

    dossiers.push({
      ref: safe(jsonStr.ref, 'SANS_REF'),
      client: safe(jsonStr.client, 'CLIENT'),
      data: jsonStr,
      postesDetail: postesDetail
    });
  }

  if (dossiers.length === 0) {
    return { status: 'error', message: 'Aucun dossier avec postes détaillés et dates' };
  }

  // Collecter toutes les années pour les fériés
  var allYears = {};
  dossiers.forEach(function(d) {
    d.postesDetail.forEach(function(p) {
      if (p.rdvs && p.rdvs[0] && p.rdvs[0].date) {
        var y = new Date(p.rdvs[0].date).getFullYear();
        allYears[y] = true;
      }
    });
  });
  var feriesSet = {};
  Object.keys(allYears).forEach(function(y) {
    var yf = getPlanningFeriesSet(parseInt(y));
    for (var k in yf) feriesSet[k] = true;
  });

  // Pour chaque dossier, générer la liste des jours avec ressources
  var clientsData = []; // { label, ref, jourMap: { 'YYYY-MM-DD': { effectif, vehicules, materiel, instructions } } }

  dossiers.forEach(function(d) {
    var jourMap = {};

    d.postesDetail.forEach(function(p) {
      var dateDebut = p.rdvs && p.rdvs[0] && p.rdvs[0].date ? new Date(p.rdvs[0].date) : null;
      if (!dateDebut || isNaN(dateDebut.getTime())) return;

      var nbJours = parseFloat(p.jours) || 1;
      var joursOuvres = genererJoursOuvres(dateDebut, Math.ceil(nbJours), feriesSet);

      // Extraire les ressources
      var effectifTotal = 0;
      (p.personnel || []).forEach(function(s) {
        var m = String(s).match(/^(\d+)x/);
        if (m) effectifTotal += parseInt(m[1]);
      });

      var vehiculesStr = (p.vehicules || []).join(', ');
      var enginsStr = (p.engins || []).join(', ');
      var allVehicules = [vehiculesStr, enginsStr].filter(function(s) { return s; }).join(', ');

      var materielStr = (p.materiel || []).join(', ');
      var instructions = safe(p.tache, p.titre || '');

      joursOuvres.forEach(function(jour) {
        var mm = String(jour.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
        var dd = String(jour.getDate()); if (dd.length < 2) dd = '0' + dd;
        var key = jour.getFullYear() + '-' + mm + '-' + dd;

        if (!jourMap[key]) {
          jourMap[key] = { effectif: 0, vehicules: [], materiel: [], instructions: [] };
        }
        jourMap[key].effectif += effectifTotal;
        if (allVehicules) jourMap[key].vehicules.push(allVehicules);
        if (materielStr) jourMap[key].materiel.push(materielStr);
        if (instructions) jourMap[key].instructions.push(instructions);
      });
    });

    if (Object.keys(jourMap).length > 0) {
      clientsData.push({
        label: d.client + ' (' + d.ref + ')',
        ref: d.ref,
        jourMap: jourMap
      });
    }
  });

  if (clientsData.length === 0) {
    return { status: 'error', message: 'Aucun poste avec dates valides' };
  }

  // Déterminer la plage de dates (min-max)
  var allDates = {};
  clientsData.forEach(function(c) {
    for (var k in c.jourMap) allDates[k] = true;
  });
  var dateKeys = Object.keys(allDates).sort();
  var minDate = new Date(dateKeys[0]);
  var maxDate = new Date(dateKeys[dateKeys.length - 1]);

  // Générer toutes les dates ouvrées entre min et max (inclure aussi les fériés pour les afficher)
  var allRows = [];
  var current = new Date(minDate);
  while (current <= maxDate) {
    var dow = current.getDay();
    if (dow !== 0 && dow !== 6) { // Seulement lun-ven
      allRows.push(new Date(current));
    }
    current.setDate(current.getDate() + 1);
  }

  // Couleurs par client
  var clientColors = [
    '#DBEAFE', // bleu clair
    '#DCFCE7', // vert clair
    '#FEF3C7', // jaune clair
    '#F3E8FF', // violet clair
    '#FFE4E6', // rose clair
    '#E0F2FE', // cyan clair
    '#FEF9C3', // lime clair
    '#FFEDD5'  // orange clair
  ];
  var headerColors = [
    '#3B82F6', // bleu
    '#16A34A', // vert
    '#F59E0B', // jaune
    '#8B5CF6', // violet
    '#EF4444', // rose
    '#06B6D4', // cyan
    '#84CC16', // lime
    '#F97316'  // orange
  ];

  // NOMS DE JOURS
  var nomJours = ['Dimanche', 'Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi'];

  // Créer le Google Sheet
  var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
  var user = getUserById(userId);
  var userName = user ? (user.prenom + ' ' + user.nom) : userId;
  var ssName = 'Planning_Consolidé_' + userName.replace(/\s+/g, '_');

  // Chercher si le fichier existe déjà
  var existingFiles = folder.getFilesByName(ssName);
  var planSS;
  if (existingFiles.hasNext()) {
    planSS = SpreadsheetApp.open(existingFiles.next());
  } else {
    planSS = SpreadsheetApp.create(ssName);
    // Déplacer dans le bon dossier
    var file = DriveApp.getFileById(planSS.getId());
    folder.addFile(file);
    var parents = file.getParents();
    while (parents.hasNext()) {
      var parent = parents.next();
      if (parent.getId() !== folder.getId()) parent.removeFile(file);
    }
  }

  // Utiliser la première feuille ou en créer une
  var planSheet = planSS.getSheets()[0];
  planSheet.setName('Planning Consolidé');
  planSheet.clear();

  // ---- CONSTRUIRE LES DONNÉES ----
  var nbClients = clientsData.length;
  var nbCols = 3 + (nbClients * 4) + 2; // A-C + 4 par client + 2 totaux

  // En-tête ligne 1 : fusion par client
  var header1 = ['Date', 'Jour', 'Férié'];
  for (var ci = 0; ci < nbClients; ci++) {
    header1.push(clientsData[ci].label);
    header1.push('');
    header1.push('');
    header1.push('');
  }
  header1.push('TOTAL');
  header1.push('');

  // En-tête ligne 2 : sous-colonnes
  var header2 = ['', '', ''];
  for (var ci = 0; ci < nbClients; ci++) {
    header2.push('Effectif');
    header2.push('Véhicules');
    header2.push('Matériel');
    header2.push('Activités');
  }
  header2.push('Hommes');
  header2.push('Véhicules');

  // Lignes de données
  var dataRows = [];
  allRows.forEach(function(date) {
    var mm = String(date.getMonth() + 1); if (mm.length < 2) mm = '0' + mm;
    var dd = String(date.getDate()); if (dd.length < 2) dd = '0' + dd;
    var key = date.getFullYear() + '-' + mm + '-' + dd;
    var dateFmt = dd + '.' + mm + '.' + date.getFullYear();
    var jourNom = nomJours[date.getDay()];
    var ferieName = getNomFerie(key, date.getFullYear());

    var row = [dateFmt, jourNom, ferieName];
    var totalHommes = 0;
    var totalVehiculesList = [];

    for (var ci = 0; ci < nbClients; ci++) {
      var cd = clientsData[ci].jourMap[key];
      if (cd) {
        row.push(cd.effectif || '');
        row.push(cd.vehicules.join('\n'));
        row.push(cd.materiel.join('\n'));
        row.push(cd.instructions.join('\n'));
        totalHommes += (cd.effectif || 0);
        if (cd.vehicules.length > 0) totalVehiculesList.push(cd.vehicules.join(', '));
      } else {
        row.push('');
        row.push('');
        row.push('');
        row.push('');
      }
    }

    row.push(totalHommes || '');
    row.push(totalVehiculesList.join('\n'));
    dataRows.push(row);
  });

  // Écrire tout d'un coup
  var allData = [header1, header2].concat(dataRows);
  if (allData.length > 0 && allData[0].length > 0) {
    planSheet.getRange(1, 1, allData.length, nbCols).setValues(allData);
  }

  // ---- FORMATAGE ----

  // Largeurs de colonnes
  planSheet.setColumnWidth(1, 100); // Date
  planSheet.setColumnWidth(2, 85);  // Jour
  planSheet.setColumnWidth(3, 140); // Férié
  for (var ci = 0; ci < nbClients; ci++) {
    var baseCol = 4 + ci * 4;
    planSheet.setColumnWidth(baseCol, 65);     // Effectif
    planSheet.setColumnWidth(baseCol + 1, 160); // Véhicules
    planSheet.setColumnWidth(baseCol + 2, 140); // Matériel
    planSheet.setColumnWidth(baseCol + 3, 180); // Activités
  }
  var totalCol1 = 4 + nbClients * 4;
  planSheet.setColumnWidth(totalCol1, 70);
  planSheet.setColumnWidth(totalCol1 + 1, 160);

  // Style en-tête ligne 1
  var h1Range = planSheet.getRange(1, 1, 1, nbCols);
  h1Range.setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center')
    .setBackground('#1E293B').setFontColor('white');

  // Fusion en-têtes clients (ligne 1)
  for (var ci = 0; ci < nbClients; ci++) {
    var baseCol = 4 + ci * 4;
    planSheet.getRange(1, baseCol, 1, 4).merge()
      .setBackground(headerColors[ci % headerColors.length]).setFontColor('white');
  }
  // Fusion TOTAL
  planSheet.getRange(1, totalCol1, 1, 2).merge().setBackground('#1E293B').setFontColor('white');

  // Style en-tête ligne 2
  var h2Range = planSheet.getRange(2, 1, 1, nbCols);
  h2Range.setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center')
    .setBackground('#E2E8F0').setFontColor('#334155');

  // Couleurs de fond par client (lignes de données)
  if (dataRows.length > 0) {
    for (var ci = 0; ci < nbClients; ci++) {
      var baseCol = 4 + ci * 4;
      var color = clientColors[ci % clientColors.length];
      planSheet.getRange(3, baseCol, dataRows.length, 4).setBackground(color);
    }

    // Colonnes TOTAL en gris clair
    planSheet.getRange(3, totalCol1, dataRows.length, 2).setBackground('#F1F5F9')
      .setFontWeight('bold');

    // Lignes fériées en rouge clair
    for (var r = 0; r < dataRows.length; r++) {
      if (dataRows[r][2]) { // colonne Férié non vide
        planSheet.getRange(3 + r, 1, 1, nbCols).setBackground('#FECACA')
          .setFontColor('#991B1B');
      }
    }

    // Format général données
    var dataRange = planSheet.getRange(3, 1, dataRows.length, nbCols);
    dataRange.setFontSize(9).setVerticalAlignment('top').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

    // Colonne Effectif centrée
    for (var ci = 0; ci < nbClients; ci++) {
      var baseCol = 4 + ci * 4;
      planSheet.getRange(3, baseCol, dataRows.length, 1).setHorizontalAlignment('center');
    }
    planSheet.getRange(3, totalCol1, dataRows.length, 1).setHorizontalAlignment('center');
  }

  // Figer les 2 premières lignes et 3 premières colonnes
  planSheet.setFrozenRows(2);
  planSheet.setFrozenColumns(3);

  // Bordures
  planSheet.getRange(1, 1, allData.length, nbCols)
    .setBorder(true, true, true, true, true, true, '#CBD5E1', SpreadsheetApp.BorderStyle.SOLID);

  Logger.log('Planning généré: ' + planSS.getUrl());

  return {
    status: 'success',
    url: planSS.getUrl(),
    nbClients: nbClients,
    nbJours: dataRows.length
  };
}

// ============================================
// TRAITEMENT PRINCIPAL (POST)
// ============================================
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // --- Routage actions utilisateur (pas de lock nécessaire) ---
    if (data._action === 'user_add') {
      try {
        var newUser = addUser(data);
        if (newUser) {
          return jsonResponse({ status: 'success', data: newUser });
        }
        return jsonResponse({ status: 'error', message: 'Nom et prenom requis ou doublon' });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Erreur creation: ' + err });
      }
    }

    if (data._action === 'user_upload_signature') {
      try {
        var sigUserId = data.userId || '';
        var sigData = data.data || '';
        if (!sigUserId || !sigData) return jsonResponse({ status: 'error', message: 'userId et data requis' });
        var sigId = updateUserSignature(sigUserId, sigData);
        if (sigId) {
          return jsonResponse({ status: 'success', signatureId: sigId });
        }
        return jsonResponse({ status: 'error', message: 'Utilisateur introuvable' });
      } catch (err) {
        return jsonResponse({ status: 'error', message: 'Erreur upload signature: ' + err });
      }
    }

    // --- Traitement devis standard (avec lock) ---
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return jsonResponse({ status: 'error', message: 'Serveur occupé, réessayez dans quelques secondes.' });
    }

    // Validation
    const errors = validateData(data);
    if (errors.length > 0) {
      return jsonResponse({ status: 'error', message: 'Validation: ' + errors.join(', ') });
    }

    const ref = safe(data.ref, 'SANS_REF').trim();
    const client = safe(data.client, 'CLIENT_INCONNU');
    const montantHT = parseFloat(data.montantHT) || 0;
    const estPelichet = data.typeSociete === 'Pelichet';
    const postes = safeArray(data.postes);

    // Dossier dédié à l'utilisateur (créé automatiquement si n'existe pas)
    const userId = safe(data.userId);
    const folder = getUserFolder(userId);
    const fileNameBase = 'Devis_' + ref + '_' + client.replace(/\s+/g, '_');

    // 1. GENERATION DEVIS
    Logger.log('Génération devis: ' + ref + ' - ' + client);
    const templateId = estPelichet ? CONFIG.TEMPLATE_PELICHET_ID : CONFIG.TEMPLATE_AUTRE_ID;
    const copyDevis = DriveApp.getFileById(templateId).makeCopy(fileNameBase, folder);
    const docDevis = DocumentApp.openById(copyDevis.getId());
    const bodyDevis = docDevis.getBody();

    remplirTableauPrestations(bodyDevis, postes);

    const tva = montantHT * CONFIG.TVA_RATE;
    const rplp = montantHT * CONFIG.RPLP_RATE;
    const totalTTC = montantHT + tva + rplp;

    // Recuperer les infos utilisateur (titre + signature) — userId déjà défini plus haut
    var userInfo = userId ? getUserById(userId) : null;
    var vendeurTitre = safe(data.vendeurTitre, userInfo ? userInfo.titre : '');

    const replacements = {
      '{{nom_societe}}': estPelichet ? 'PELICHET NLC SA' : safe(data.nomPrestataire, 'Société'),
      '{{ref}}': ref,
      '{{date}}': dateJour(),
      '{{salutation}}': safe(data.genre, 'Monsieur'),
      '{{vendeur}}': safe(data.vendeur, 'ARNAUD GUEDOU'),
      '{{vendeur_titre}}': vendeurTitre,
      '{{vendeur_tel}}': safe(data.vendeurTel, '+41 79 688 27 35'),
      '{{client}}': client,
      '{{adresse_client}}': safe(data.adresseClient),
      '{{contact}}': safe(data.contact),
      '{{date_prevue}}': calculerDatePrevue(data),
      '{{adresse_depart}}': safe(data.adresseDepart),
      '{{adresse_arrivee}}': safe(data.adresseArrivee),
      '{{ht}}': fmt(montantHT),
      '{{tva}}': fmt(tva),
      '{{rplp}}': fmt(rplp),
      '{{total}}': fmt(totalTTC)
    };

    for (const key in replacements) {
      bodyDevis.replaceText(key.replace(/[{}]/g, '\\$&'), replacements[key]);
    }

    // Inserer l'image de signature si disponible
    var signatureId = userInfo ? userInfo.signatureId : '';
    if (signatureId) {
      try {
        insererSignature(bodyDevis, signatureId);
      } catch (sigErr) {
        Logger.log('Erreur insertion signature: ' + sigErr);
      }
    } else {
      // Supprimer le placeholder {{signature}} s'il existe mais pas de signature
      bodyDevis.replaceText('\\{\\{signature\\}\\}', '');
    }

    docDevis.saveAndClose();

    // 2. FICHE RESA
    var resaStatus = 'skipped';
    var resaError = '';
    var resaFileId = null;
    try {
      resaFileId = genererFicheResa(data, folder, ref, client);
      resaStatus = 'ok';
      Logger.log('Fiche RESA generee: ' + ref);
    } catch (err) {
      resaStatus = 'error';
      resaError = err.toString();
      Logger.log('ERREUR RESA: ' + err + '\n' + (err.stack || ''));
    }

    // 3. PDF DEVIS + ARCHIVAGE
    const pdfBlob = copyDevis.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName(fileNameBase + '.pdf');

    // 4. PDF RESA + DOCX RESA
    var resaPdfB64 = '';
    var resaPdfName = '';
    var resaDocxB64 = '';
    var resaDocxName = '';
    if (resaFileId) {
      try {
        var resaFile = DriveApp.getFileById(resaFileId);
        var resaPdfBlob = resaFile.getAs('application/pdf');
        resaPdfB64 = Utilities.base64Encode(resaPdfBlob.getBytes());
        resaPdfName = resaFile.getName() + '.pdf';

        // Export DOCX via API Drive
        var resaDocxBlob = exportDocAsDocx(resaFileId);
        if (resaDocxBlob) {
          resaDocxB64 = Utilities.base64Encode(resaDocxBlob.getBytes());
          resaDocxName = resaFile.getName() + '.docx';
        }
      } catch (e) {
        Logger.log('Erreur export RESA: ' + e);
      }
    }

    // 5. DOCX DEVIS
    var docxB64 = '';
    var docxName = '';
    try {
      var devisDocxBlob = exportDocAsDocx(copyDevis.getId());
      if (devisDocxBlob) {
        docxB64 = Utilities.base64Encode(devisDocxBlob.getBytes());
        docxName = fileNameBase + '.docx';
      }
    } catch (e) {
      Logger.log('Erreur export DOCX devis: ' + e);
    }

    // Encoder le PDF devis en base64
    var pdfB64 = Utilities.base64Encode(pdfBlob.getBytes());

    archiver(ref, estPelichet, client, data, montantHT, pdfFile);

    Logger.log('Devis terminé: ' + ref + ' -> ' + pdfFile.getUrl());

    return jsonResponse({
      status: 'success',
      pdfUrl: pdfFile.getUrl(),
      docUrl: copyDevis.getUrl(),
      resa: resaStatus,
      resaError: resaError,
      pdfB64: pdfB64,
      pdfName: fileNameBase + '.pdf',
      docxB64: docxB64,
      docxName: docxName,
      resaPdfB64: resaPdfB64,
      resaPdfName: resaPdfName,
      resaDocxB64: resaDocxB64,
      resaDocxName: resaDocxName
    });

  } catch (error) {
    Logger.log('ERREUR CRITIQUE doPost: ' + error + '\n' + error.stack);
    return jsonResponse({ status: 'error', message: error.toString() });
  } finally {
    if (lock) lock.releaseLock();
  }
}

// ============================================
// ARCHIVAGE SPREADSHEET
// ============================================
function archiver(ref, estPelichet, client, data, montantHT, pdfFile) {
  try {
    const ss = getSs();
    if (!ss) return;

    // Determiner l'onglet : si userId present, utiliser l'onglet user (le creer si besoin)
    var userId = safe(data.userId);
    var sheet;
    if (userId) {
      sheet = ensureUserSheet(userId);
    } else {
      sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
      if (!sheet) {
        var sheets = ss.getSheets();
        sheet = sheets.length > 0 ? sheets[0] : ss.insertSheet(CONFIG.SHEET_NAME);
      }
    }

    const tva = montantHT * CONFIG.TVA_RATE;
    const rplp = montantHT * CONFIG.RPLP_RATE;
    const totalTTC = montantHT + tva + rplp;

    // Résumé des postes pour la colonne "Détail Tarifié"
    // Format existant : [TITRE: 1 500,00 CHF] [TITRE2: 500,00 CHF]
    var detailTarifie = safeArray(data.postes).map(function(p) {
      return '[' + safe(p.titre, 'PRESTATION').toUpperCase() + ': ' + fmt(parseFloat(p.prix) || 0) + ']';
    }).join(' ');

    // Colonnes alignées sur le Sheet existant :
    // Date | Réf | Société | Email | Statut Email | Client | Facturation | Contact | Départ | Arrivée | Prévu | Détail Tarifié | HT Global | TTC Global | Lien | JSON
    sheet.appendRow([
      new Date(),                           // A: Date
      ref,                                  // B: Réf
      estPelichet ? 'Pelichet' : 'Autre',   // C: Société
      safe(data.email),                      // D: Email
      'OK',                                  // E: Statut Email
      client,                                // F: Client
      safe(data.adresseClient),              // G: Facturation
      safe(data.contact),                    // H: Contact
      safe(data.adresseDepart),              // I: Départ
      safe(data.adresseArrivee),             // J: Arrivée
      safe(data.datePrevue),                 // K: Prévu
      detailTarifie,                         // L: Détail Tarifié
      montantHT,                             // M: HT Global
      totalTTC,                              // N: TTC Global
      pdfFile.getUrl(),                      // O: Lien
      JSON.stringify(data)                   // P: JSON (rechargement complet)
    ]);
  } catch (err) {
    Logger.log('ERREUR archivage (non bloquante): ' + err);
  }
}

// ============================================
// REMPLISSAGE TABLEAU PRESTATIONS
// ============================================
function remplirTableauPrestations(body, postes) {
  const tables = body.getTables();
  let prestaTable = null;
  let rowIndex = -1;

  for (let i = 0; i < tables.length; i++) {
    for (let j = 0; j < tables[i].getNumRows(); j++) {
      if (tables[i].getRow(j).getText().indexOf('{{detail_prestations}}') !== -1) {
        prestaTable = tables[i];
        rowIndex = j;
        break;
      }
    }
    if (prestaTable) break;
  }

  if (!prestaTable || rowIndex === -1 || !postes || postes.length === 0) return;

  const templateRow = prestaTable.getRow(rowIndex);

  postes.forEach(function(p, idx) {
    const newRow = prestaTable.insertTableRow(rowIndex + idx + 1, templateRow.copy());
    const titre = safe(p.titre, 'PRESTATION').toUpperCase();
    const nbJ = parseFloat(p.jours) || 1;
    const joursFmt = nbJ === 0.5 ? '1/2 jour' : (nbJ + (nbJ > 1 ? ' jours' : ' jour'));

    let pDesc;
    if (p.mode === 'simple') {
      pDesc = titre + '\n' + safe(p.text);
    } else {
      const personnel = transformerPersonnel(safeArray(p.personnel)).join(', ');
      const vehicules = safeArray(p.vehicules).join(', ');
      pDesc = titre + '\n' + safe(p.tache) +
        '\n\u2022 Durée : ' + joursFmt +
        '\n\u2022 Ressources : ' + personnel + ' / ' + vehicules;
    }

    newRow.getCell(0).setText(pDesc.trim());
    // BUG FIX: parseFloat avant fmt pour éviter "NaN CHF"
    newRow.getCell(1).setText(p.prix ? fmt(parseFloat(p.prix)) : 'Inclus');

    // Mettre en gras le titre + supprimer tout surlignage hérité du template
    try {
      var txt0 = newRow.getCell(0).getChild(0).asParagraph().editAsText();
      txt0.setBold(0, titre.length, true);
      // Supprimer le surlignage sur toute la cellule
      var fullText = txt0.getText();
      if (fullText.length > 0) {
        txt0.setBackgroundColor(0, fullText.length - 1, null);
      }
      var txt1 = newRow.getCell(1).getChild(0).asParagraph().editAsText();
      var t1 = txt1.getText();
      if (t1.length > 0) {
        txt1.setBackgroundColor(0, t1.length - 1, null);
      }
    } catch (e) { /* silencieux si mise en forme échoue */ }
  });

  prestaTable.removeRow(rowIndex);
}

// ============================================
// GENERATION FICHE RESA
// ============================================
function genererFicheResa(data, folder, ref, client) {
  const fileNameResa = ref + ' RESA - ' + client.replace(/\s+/g, '_');
  const copyResa = DriveApp.getFileById(CONFIG.TEMPLATE_RESA_ID).makeCopy(fileNameResa, folder);
  const resaFileId = copyResa.getId();
  const docResa = DocumentApp.openById(resaFileId);
  const body = docResa.getBody();

  const volumeEstime = safe(data.volumeEstime, '200');

  // Remplacement des placeholders (double escaping pour replaceText regex)
  var replacements = {
    'ref': ref,
    'coordinateur': safe(data.coordinateur, 'WOTQUENNE LILOU').toUpperCase(),
    'coordinateur_tel': safe(data.coordinateurTel, '022 827 36 97'),
    'vendeur': safe(data.vendeur, 'ARNAUD GUEDOU').toUpperCase(),
    'vendeur_tel': safe(data.vendeurTel, '+41 79 688 27 35'),
    'client': client,
    'adresse_depart': safe(data.adresseDepart),
    'adresse_arrivee': safe(data.adresseArrivee),
    'volume': volumeEstime + ' M3',
    'contact': safe(data.contact),
    'tel_client': [safe(data.contact), safe(data.contactTel)].filter(function(s) { return s; }).join(' – ')
  };

  for (var key in replacements) {
    // Matcher {{key}} avec regex escaping
    var pattern = '\\{\\{' + key + '\\}\\}';
    try {
      body.replaceText(pattern, replacements[key]);
    } catch (e) {
      Logger.log('RESA replaceText failed for ' + key + ': ' + e);
    }
  }

  // ---- TABLEAU RESSOURCES ----
  // Trouver la table de log : on essaye plusieurs mots-clés en-tête
  // (date, heure, hommes, vehicule, ressource) avec accents/sans, casse ignorée
  var tables = body.getTables();
  var logTable = null;
  var headerKeywords = ['date', 'heure', 'homme', 'vehicule', 'véhicule', 'ressource', 'effectif'];
  Logger.log('RESA: ' + tables.length + ' tables trouvees');
  for (var t = 0; t < tables.length; t++) {
    try {
      var headerText = tables[t].getRow(0).getText().toLowerCase();
      Logger.log('RESA table[' + t + '] header: ' + headerText.substring(0, 100));
      for (var hk = 0; hk < headerKeywords.length; hk++) {
        if (headerText.indexOf(headerKeywords[hk]) !== -1) {
          logTable = tables[t];
          Logger.log('RESA: log table = table[' + t + '] (matched keyword: ' + headerKeywords[hk] + ')');
          break;
        }
      }
      if (logTable) break;
    } catch (e) { /* ignore */ }
  }
  // Fallback : si aucune table avec ces mots-clés, prendre la dernière table du document
  // (souvent c'est celle des ressources, vide attendant des lignes)
  if (!logTable && tables.length > 0) {
    logTable = tables[tables.length - 1];
    Logger.log('RESA: fallback - using last table as log table');
  }

  // ---- SUPPRIMER "PRESTATION X" du template si present ----
  var prestaRange = body.findText('PRESTATION\\s*\\d*');
  while (prestaRange) {
    var prestaEl = prestaRange.getElement();
    var prestaPara = prestaEl.getParent();
    try {
      // Supprimer seulement si le paragraphe ne contient que "PRESTATION X"
      var paraText = prestaPara.asParagraph().getText().trim();
      if (/^PRESTATION\s*\d*$/i.test(paraText)) {
        prestaPara.removeFromParent();
      }
    } catch(e) { /* ignore */ }
    prestaRange = body.findText('PRESTATION\\s*\\d*');
  }

  // ---- INSTRUCTIONS DETAILLEES ----
  // Police et tailles standard pour toutes les lignes inserees
  var RESA_FONT = 'Trebuchet MS';
  var FONT_SIZE_RDV = 14;
  var FONT_SIZE_TITRE = 12;
  var FONT_SIZE_TEXTE = 12;
  var FONT_SIZE_RESSOURCE = 10;
  var COLOR_RDV = '#C00000'; // Rouge foncé pour les lignes RDV

  // Helper : reset complet du style d'un paragraphe pour eviter l'heritage
  function resetStyle(p, opts) {
    var t = p.editAsText();
    t.setFontFamily(RESA_FONT);
    t.setFontSize(opts.size || FONT_SIZE_TEXTE);
    t.setBold(opts.bold || false);
    t.setItalic(opts.italic || false);
    t.setUnderline(opts.underline || false);
    t.setForegroundColor(opts.color || '#000000');
    return t;
  }

  // Recherche du placeholder {{instructions}} avec plusieurs stratégies
  var instRange = body.findText('\\{\\{instructions\\}\\}');
  var paragraph = null;
  var container = null;
  var insertionIndex = -1;

  if (instRange) {
    Logger.log('RESA: placeholder {{instructions}} trouvé via findText');
    var element = instRange.getElement();
    paragraph = element.getParent().asParagraph();
    container = paragraph.getParent();
    insertionIndex = container.getChildIndex(paragraph);
  } else {
    // Fallback 1 : chercher par itération paragraph par paragraph (gère les runs splités)
    Logger.log('RESA: findText {{instructions}} a échoué, tentative paragraphes');
    var paras = body.getParagraphs();
    for (var pi = 0; pi < paras.length; pi++) {
      var ptext = paras[pi].getText();
      if (ptext && ptext.indexOf('{{instructions}}') !== -1) {
        paragraph = paras[pi];
        container = paragraph.getParent();
        insertionIndex = container.getChildIndex(paragraph);
        Logger.log('RESA: placeholder trouvé via getParagraphs au paragraphe ' + pi);
        break;
      }
      // Match aussi sans accolades cassées (ex: {{ instructions }})
      if (ptext && /\{\s*\{\s*instructions\s*\}\s*\}/i.test(ptext)) {
        paragraph = paras[pi];
        container = paragraph.getParent();
        insertionIndex = container.getChildIndex(paragraph);
        Logger.log('RESA: placeholder trouvé via regex tolérant');
        break;
      }
    }
  }

  if (paragraph && container && insertionIndex >= 0) {
    var allPostes = safeArray(data.postes);

    // Traiter TOUS les postes (simples et détaillés)
    allPostes.forEach(function(p, posteIdx) {
      var titre = safe(p.titre, '').toUpperCase();
      if (!titre) return; // Pas de titre = pas d'insertion
      var nbJoursVal = parseFloat(p.jours) || 1;
      var joursFmt = nbJoursVal === 0.5 ? '1/2 jour' : (nbJoursVal + (nbJoursVal > 1 ? ' jours' : ' jour'));

      if (p.mode === 'simple') {
        // ---- POSTE SIMPLE : titre + description ----
        var pTitleS = container.insertParagraph(insertionIndex++, titre);
        resetStyle(pTitleS, { bold: true, underline: true, size: FONT_SIZE_TITRE });

        if (safe(p.text)) {
          var pTextS = container.insertParagraph(insertionIndex++, safe(p.text));
          resetStyle(pTextS, { size: FONT_SIZE_TEXTE });
        }

        // Ligne vide de séparation
        var pEmpty = container.insertParagraph(insertionIndex++, '');
        resetStyle(pEmpty, { size: FONT_SIZE_TEXTE });

      } else {
        // ---- POSTE DETAIL : RDV + ressources ----

        // Collecter les dates RDV saisies par l'utilisateur
        var rdvs = safeArray(p.rdvs).filter(function(r) { return r.date; });

        // Si aucun RDV saisi, en créer un par défaut
        if (rdvs.length === 0) {
          rdvs = [{ date: new Date(), heure: '8H00' }];
        }

        // Afficher les RDV (uniquement ceux saisis, sans en générer d'extras)
        var rdvLines = [];
        rdvs.forEach(function(r) {
          var d = (r.date instanceof Date) ? r.date : new Date(r.date);
          var dateFmt = Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          rdvLines.push(dateFmt + ' a ' + safe(r.heure, '8H00'));
        });

        // Ligne RDV — rouge vif, Trebuchet MS 12pt
        var rdvText = 'RDV : ' + rdvLines.join(' / ') + ' (' + joursFmt + ')';
        var pRDV = container.insertParagraph(insertionIndex++, rdvText);
        resetStyle(pRDV, { bold: true, underline: true, size: FONT_SIZE_RDV, color: COLOR_RDV });

        // Titre du poste — noir, gras souligné (directement après RDV, sans ligne vide)
        var pTitle = container.insertParagraph(insertionIndex++, titre);
        resetStyle(pTitle, { bold: true, underline: true, size: FONT_SIZE_TITRE });

        // Tache / instructions — normal, pas d'italic ni underline
        if (safe(p.tache)) {
          var pTask = container.insertParagraph(insertionIndex++, safe(p.tache));
          resetStyle(pTask, { size: FONT_SIZE_TEXTE });
        }

        // Uniquement le matériel dans la section texte
        // (personnel, véhicules et engins sont deja dans le tableau ressources)
        var materielList = safeArray(p.materiel);
        Logger.log('RESA materiel raw: ' + JSON.stringify(p.materiel) + ' => list: ' + JSON.stringify(materielList));
        if (materielList.length > 0) {
          // Ligne vide entre tache et matériel
          var pEmptyMat = container.insertParagraph(insertionIndex++, '');
          resetStyle(pEmptyMat, { size: FONT_SIZE_TEXTE });

          var pMat = container.insertParagraph(insertionIndex++, 'Matériel : ' + materielList.join(', '));
          resetStyle(pMat, { italic: true, size: FONT_SIZE_RESSOURCE });
        }

        // Ligne vide de séparation — reset complet pour couper l'heritage
        var pEmptyD = container.insertParagraph(insertionIndex++, '');
        resetStyle(pEmptyD, { size: FONT_SIZE_TEXTE });

        // ---- Tableau ressources : 1 ligne par poste (plage de dates) ----
        if (logTable) {
          // Date de debut = premier RDV saisi (normaliser a midi pour eviter timezone)
          var rawDate = (rdvs[0].date instanceof Date) ? rdvs[0].date : new Date(rdvs[0].date);
          var dateDebut = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate(), 12, 0, 0);
          var heureRDV = safe(rdvs[0].heure, '8H00');

          // Personnel et vehicules pour le tableau
          var personnelList = transformerPersonnel(safeArray(p.personnel));
          var vehiculesList = safeArray(p.vehicules);
          var enginsList = safeArray(p.engins);

          // Generer les jours ouvres pour trouver la date de fin
          var joursOuvres = getJoursOuvres(dateDebut, nbJoursVal);
          var dateFin = joursOuvres[joursOuvres.length - 1];

          // Format de la colonne Date
          var dateDebutFmt = Utilities.formatDate(dateDebut, CONFIG.TIMEZONE, 'dd.MM.yyyy');
          var dateCell;
          if (nbJoursVal <= 1) {
            // 1 jour ou demi-jour : juste la date
            dateCell = dateDebutFmt;
          } else {
            // Plusieurs jours : plage "debut au fin"
            var dateFinFmt = Utilities.formatDate(dateFin, CONFIG.TIMEZONE, 'dd.MM.yyyy');
            dateCell = dateDebutFmt + ' au ' + dateFinFmt;
          }

          // Format de la durée
          var dureeCell = nbJoursVal === 0.5 ? '1/2 jour' : (nbJoursVal + (nbJoursVal > 1 ? ' jours' : ' jour'));

          var row = logTable.appendTableRow();

          // appendTableRow() cree parfois une cellule par defaut — la supprimer
          while (row.getNumCells() > 0) {
            row.removeCell(0);
          }

          var cellValues = [dateCell, heureRDV, personnelList.join('\n'), dureeCell, vehiculesList.join('\n'), enginsList.join('\n')];
          for (var ci = 0; ci < cellValues.length; ci++) {
            var cell = row.appendTableCell(cellValues[ci]);
            // Forcer fond blanc et style propre
            cell.setBackgroundColor('#FFFFFF');
            // Styler le texte de chaque paragraphe dans la cellule
            for (var pi = 0; pi < cell.getNumChildren(); pi++) {
              var child = cell.getChild(pi);
              if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
                var txt = child.editAsText();
                txt.setFontFamily(RESA_FONT);
                txt.setFontSize(9);
                txt.setBold(false);
                txt.setItalic(false);
                txt.setUnderline(false);
                txt.setForegroundColor('#000000');
                txt.setBackgroundColor(null);
              }
            }
          }
        }
      }
    });

    // Supprimer le paragraphe {{instructions}}
    try { paragraph.removeFromParent(); } catch (e) { /* ignore */ }
  } else {
    Logger.log('RESA ERREUR: placeholder {{instructions}} introuvable dans le template. ' +
               'Vérifie que le doc template (ID: ' + CONFIG.TEMPLATE_RESA_ID + ') contient bien {{instructions}} ' +
               'sur une ligne dédiée. Astuce : retape le placeholder en une seule fois (sans copier-coller).');
  }

  docResa.saveAndClose();
  return resaFileId;
}

// ============================================
// INSERTION SIGNATURE DANS UN DOCUMENT
// ============================================

/**
 * Insere l'image de signature a la place du placeholder {{signature}}
 * La signature est redimensionnee a ~150px de large
 */
function insererSignature(body, signatureFileId) {
  if (!signatureFileId) return;

  var sigRange = body.findText('\\{\\{signature\\}\\}');
  if (!sigRange) {
    Logger.log('Placeholder {{signature}} non trouve dans le document');
    return;
  }

  var element = sigRange.getElement();
  var paragraph = element.getParent().asParagraph();

  // Recuperer l'image depuis Drive
  var sigFile = DriveApp.getFileById(signatureFileId);
  var sigBlob = sigFile.getBlob();

  // Supprimer le texte placeholder
  paragraph.clear();

  // Inserer l'image
  var img = paragraph.appendInlineImage(sigBlob);

  // Redimensionner (largeur max 200px, hauteur proportionnelle)
  var origW = img.getWidth();
  var origH = img.getHeight();
  var maxW = 200;
  if (origW > maxW) {
    var ratio = maxW / origW;
    img.setWidth(maxW);
    img.setHeight(Math.round(origH * ratio));
  }

  Logger.log('Signature inseree (ID: ' + signatureFileId + ')');
}

// ============================================
// REPONSE JSON
// ============================================
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================
// INITIALISATION DES TEMPLATES (exécuter manuellement une seule fois)
// ============================================
function creerModeleReservation() {
  const doc = DocumentApp.create('TEMPLATE_RESA_V2');
  const body = doc.getBody();
  const header = body.appendTable([
    ['Dossier : {{ref}}', 'Coordinateur : {{coordinateur}}', 'Tél : 022 827 36 97'],
    ['', 'Vendeur : {{vendeur}}', 'Tél : {{vendeur_tel}}']
  ]);
  header.setBorderWidth(0);
  body.appendParagraph('\nENLÈVEMENT / LIVRAISON').setBold(true).setUnderline(true);
  body.appendTable([
    ['DÉPART :\n{{client}}\n{{adresse_depart}}\n\nTel client : {{tel_client}}',
     'ARRIVÉE :\n{{client}}\n{{adresse_arrivee}}']
  ]);
  body.appendParagraph('\nVOLUME : {{volume}}').setBold(true);
  body.appendParagraph('\nINSTRUCTIONS :\n{{instructions}}');
  body.appendParagraph('\nRESSOURCES :');
  const logTable = body.appendTable([['Date', 'Heure', 'Hommes', 'Durée', 'Véhicule', 'Véhicules Spéciaux']]);
  logTable.getRow(0).setBackgroundColor('#EEEEEE').setBold(true);
  Logger.log('ID RESA : ' + doc.getId());
}

function creerModeleDevis() {
  const doc = DocumentApp.create('TEMPLATE_DEVIS_PELICHET_V2');
  const body = doc.getBody();
  body.appendParagraph('PELICHET NLC SA').setBold(true).setFontSize(14)
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('Conseiller : {{vendeur}}\nTél : {{vendeur_tel}}\nRéférence : {{ref}}')
    .setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
  body.appendParagraph('\n{{salutation}}\n{{client}}\n{{adresse_client}}').setBold(true);
  body.appendParagraph('\nOBJET : DEVIS LOGISTIQUE').setBold(true).setUnderline(true);
  body.appendParagraph('Départ : {{adresse_depart}}\nArrivée : {{adresse_arrivee}}');
  const table = body.appendTable([['Description', 'Montant HT']]);
  table.getRow(0).setBackgroundColor(CONFIG.COLOR_PELICHET).setBold(true);
  table.appendTableRow().appendTableCell('{{detail_prestations}}')
    .getParentRow().appendTableCell('0.00 CHF');
  body.appendParagraph(
    '\nTOTAL HT : {{ht}}\nTVA (8.1%) : {{tva}}\nRPLP (0.5%) : {{rplp}}\nTOTAL TTC : {{total}}'
  ).setBold(true);
  Logger.log('ID DEVIS : ' + doc.getId());
}
