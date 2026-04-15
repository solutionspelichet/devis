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
  TEMPLATE_PELICHET_ID: '17LFpCQWcyiYgvgAM9Og_wkzsu1GwK2hmYZFICSHfSxg',
  TEMPLATE_AUTRE_ID: '1QWDp-ACk1dL7bXF2fq5VQRbPW4jxwXtUmq8hA1z4cD0',
  TEMPLATE_RESA_ID: '15QenLctlfoeBb--sC8WRM-Mi5NJjUn8kr01CMsnrwQg',
  FOLDER_ID: '1DKVN4T2gKhPf26qOpT0sGiXrWHzSUBDk',
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
  var next = new Date(date.getTime());
  next.setDate(next.getDate() + 1);
  while (!isJourOuvre(next)) {
    next.setDate(next.getDate() + 1);
  }
  return next;
}

/**
 * Genere la liste des jours ouvres a partir d'une date de debut pour N jours
 * Ex: getJoursOuvres(date, 10) -> tableau de 10 dates de jours ouvres
 */
function getJoursOuvres(dateDebut, nbJours) {
  var result = [];
  var current = new Date(dateDebut.getTime());

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
  var folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);

  // Chercher ou creer le sous-dossier Signatures
  var sigFolders = folder.getFoldersByName(CONFIG.SIGNATURES_FOLDER_NAME);
  var sigFolder;
  if (sigFolders.hasNext()) {
    sigFolder = sigFolders.next();
  } else {
    sigFolder = folder.createFolder(CONFIG.SIGNATURES_FOLDER_NAME);
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
    var hasDemiJour = (nbJoursVal % 1 !== 0);
    var titre = safe(p.titre, 'PRESTATION').toUpperCase();

    // --- PERSONNEL : 1 ligne par jour ---
    var personnelItems = safeArray(p.personnel);
    personnelItems.forEach(function(item) {
      var match = String(item).match(/^(\d+)x\s+(.+)$/);
      if (!match) return;
      var qty = parseInt(match[1]) || 1;
      var nom = match[2].trim();

      // Chercher le cout unitaire dans les tarifs
      var coutUnit = 0;
      if (tarifs.items.personnel) {
        var found = tarifs.items.personnel.find(function(t) { return t.item === nom; });
        if (found) coutUnit = parseFloat(found.cout) || 0;
      }

      joursOuvres.forEach(function(jourDate, idx) {
        var dateFmt = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
        var isLast = (idx === joursOuvres.length - 1);
        var duree = (isLast && hasDemiJour) ? '1/2J' : '1J';
        var multiplicateur = (isLast && hasDemiJour) ? 0.5 : 1;
        var montant = coutUnit * qty * multiplicateur;
        var abrNom = nom.replace(/Manutentionnaire du lourd/gi, 'HL')
                        .replace(/Manutentionnaire/gi, 'H')
                        .replace(/Chauffeur/gi, 'CHAUF')
                        .replace(/Chef d'equipe/gi, 'CE');

        var posteCode = (isLast && hasDemiJour) ? V.personnel_demi : V.personnel;

        lignes.push({
          n: n++,
          ventil: posteCode.ventil,
          poste: posteCode.poste,
          libelle: dateFmt + ' - ' + qty + ' ' + abrNom + ' - ' + duree + ' - ' + titre,
          montant: montant,
          dev: 'CHF'
        });
      });
    });

    // --- VEHICULES : 1 ligne par jour ---
    var vehiculesItems = safeArray(p.vehicules);
    vehiculesItems.forEach(function(item) {
      var match = String(item).match(/^(\d+)x\s+(.+)$/);
      if (!match) return;
      var qty = parseInt(match[1]) || 1;
      var nom = match[2].trim();

      var coutUnit = 0;
      if (tarifs.items.vehicules) {
        var found = tarifs.items.vehicules.find(function(t) { return t.item === nom; });
        if (found) coutUnit = parseFloat(found.cout) || 0;
      }

      joursOuvres.forEach(function(jourDate, idx) {
        var dateFmt = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
        var isLast = (idx === joursOuvres.length - 1);
        var duree = (isLast && hasDemiJour) ? '1/2J' : '1J';
        var multiplicateur = (isLast && hasDemiJour) ? 0.5 : 1;
        var montant = coutUnit * qty * multiplicateur;

        lignes.push({
          n: n++,
          ventil: V.vehicules.ventil,
          poste: V.vehicules.poste,
          libelle: dateFmt + ' - ' + qty + ' ' + nom + ' - ' + duree + ' - ' + titre,
          montant: montant,
          dev: 'CHF'
        });
      });
    });

    // --- ENGINS : 1 ligne par jour ---
    var enginsItems = safeArray(p.engins);
    enginsItems.forEach(function(item) {
      var match = String(item).match(/^(\d+)x\s+(.+?)(?:\s+\(\d+T\))?$/);
      if (!match) return;
      var qty = parseInt(match[1]) || 1;
      var nom = match[2].trim();

      var coutUnit = 0;
      if (tarifs.items.engins) {
        var found = tarifs.items.engins.find(function(t) { return t.item === nom; });
        if (found) coutUnit = parseFloat(found.cout) || 0;
      }

      var enginConf = V.engins[nom] || V.engins._default;

      joursOuvres.forEach(function(jourDate, idx) {
        var dateFmt = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
        var isLast = (idx === joursOuvres.length - 1);
        var duree = (isLast && hasDemiJour) ? '1/2J' : '1J';
        var multiplicateur = (isLast && hasDemiJour) ? 0.5 : 1;
        var montant = coutUnit * qty * multiplicateur;

        lignes.push({
          n: n++,
          ventil: enginConf.ventil,
          poste: enginConf.poste,
          libelle: dateFmt + ' - ' + qty + ' ' + nom + ' - ' + duree + ' - ' + titre,
          montant: montant,
          dev: 'CHF'
        });
      });
    });

    // --- MATERIEL : 1 seule ligne (pas par jour, c'est des pieces) ---
    var materielItems = safeArray(p.materiel);
    materielItems.forEach(function(item) {
      var match = String(item).match(/^(\d+)x\s+(.+)$/);
      if (!match) return;
      var qty = parseInt(match[1]) || 1;
      var nom = match[2].trim();

      var coutUnit = 0;
      var unite = 'piece';
      if (tarifs.items.materiel) {
        var found = tarifs.items.materiel.find(function(t) { return t.item === nom; });
        if (found) { coutUnit = parseFloat(found.cout) || 0; unite = found.unite || 'piece'; }
      }

      // Si unite = jour, multiplier par nbJours
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

  // En-tetes : Type | Item | Cout unitaire | Unite | Marge %
  for (var i = 1; i < values.length; i++) {
    var type = String(values[i][0] || '').trim();
    var item = String(values[i][1] || '').trim();
    var cout = parseFloat(values[i][2]) || 0;
    var unite = String(values[i][3] || 'jour').trim();
    var marge = parseFloat(values[i][4]);

    if (!type) continue;

    if (type === '_marge') {
      // Ligne de marge par categorie
      marges[item] = cout;
    } else {
      // Ligne de tarif
      if (!items[type]) items[type] = [];
      items[type].push({ item: item, cout: cout, unite: unite });
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
  sheet.appendRow(['Type', 'Item', 'Cout unitaire', 'Unite', 'Marge %']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');

  // Marges par categorie
  var marges = data.marges || {};
  for (var cat in marges) {
    sheet.appendRow(['_marge', cat, marges[cat], '%', '']);
  }

  // Items
  var items = data.items || {};
  for (var type in items) {
    var list = items[type] || [];
    for (var i = 0; i < list.length; i++) {
      sheet.appendRow([type, list[i].item, list[i].cout, list[i].unite || 'jour', marges[type] || '']);
    }
  }

  // Formater les colonnes
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);

  return true;
}

/**
 * Cree l'onglet Tarifs avec des valeurs par defaut
 */
function creerOngletTarifs(ss) {
  var sheet = ss.insertSheet(TARIFS_SHEET_NAME);
  sheet.appendRow(['Type', 'Item', 'Cout unitaire', 'Unite', 'Marge %']);
  sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#D32F2F').setFontColor('#FFFFFF');

  // Marges par defaut
  sheet.appendRow(['_marge', 'personnel', 35, '%', '']);
  sheet.appendRow(['_marge', 'vehicules', 25, '%', '']);
  sheet.appendRow(['_marge', 'engins', 40, '%', '']);
  sheet.appendRow(['_marge', 'materiel', 10, '%', '']);

  // Tarifs par defaut
  var defauts = [
    ['personnel', 'Manutentionnaire', 350, 'jour'],
    ['personnel', 'Manutentionnaire du lourd', 400, 'jour'],
    ['personnel', 'Chauffeur', 380, 'jour'],
    ['personnel', "Chef d'equipe", 450, 'jour'],
    ['personnel', 'CHAUF-LIVREUR', 400, 'jour'],
    ['vehicules', 'VL', 200, 'jour'],
    ['vehicules', '1F', 350, 'jour'],
    ['vehicules', 'PL', 500, 'jour'],
    ['vehicules', 'Semi', 700, 'jour'],
    ['vehicules', 'Box IT', 150, 'jour'],
    ['engins', 'Chariot elevateur', 600, 'jour'],
    ['engins', 'Grue mobile', 1200, 'jour'],
    ['engins', 'Monte-meuble', 400, 'jour'],
    ['materiel', 'Chariots', 5, 'piece'],
    ['materiel', 'Rouleaux bulle', 15, 'piece'],
    ['materiel', 'Adhesif', 3, 'piece'],
    ['materiel', 'Transpalette', 30, 'jour']
  ];
  defauts.forEach(function(row) { sheet.appendRow(row); });

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);

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
  const dossiers = [];

  for (let i = values.length - 1; i >= 1 && dossiers.length < 200; i--) {
    const ref = String(values[i][refIdx] || '').trim();
    if (!ref) continue;

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
 * Jours fériés suisses (Genève) pour les années courantes.
 * Retourne un Set de strings au format 'YYYY-MM-DD'
 */
function getJoursFeries(year) {
  var feries = [];
  // Fériés fixes
  feries.push(year + '-01-01'); // Nouvel An
  feries.push(year + '-01-02'); // 2 janvier
  feries.push(year + '-03-01'); // Instauration de la République (GE)
  feries.push(year + '-05-01'); // Fête du Travail
  feries.push(year + '-06-01'); // Lundi de Pentecôte (approximatif)
  feries.push(year + '-08-01'); // Fête nationale
  feries.push(year + '-09-11'); // Jeûne genevois (GE, variable)
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
    var yf = getJoursFeries(parseInt(y));
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

    const folder = DriveApp.getFolderById(CONFIG.FOLDER_ID);
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

    // Recuperer les infos utilisateur (titre + signature)
    var userId = safe(data.userId);
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
      '{{date_prevue}}': safe(data.datePrevue),
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
    try {
      genererFicheResa(data, folder, ref, client);
      resaStatus = 'ok';
      Logger.log('Fiche RESA generee: ' + ref);
    } catch (err) {
      resaStatus = 'error';
      resaError = err.toString();
      Logger.log('ERREUR RESA: ' + err + '\n' + (err.stack || ''));
    }

    // 3. PDF + ARCHIVAGE
    const pdfBlob = copyDevis.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName(fileNameBase + '.pdf');

    archiver(ref, estPelichet, client, data, montantHT, pdfFile);

    Logger.log('Devis terminé: ' + ref + ' -> ' + pdfFile.getUrl());

    return jsonResponse({
      status: 'success',
      pdfUrl: pdfFile.getUrl(),
      docUrl: copyDevis.getUrl(),
      resa: resaStatus,
      resaError: resaError
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

    try {
      newRow.getCell(0).getChild(0).asParagraph().editAsText().setBold(0, titre.length, true);
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
  const docResa = DocumentApp.openById(copyResa.getId());
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
    'tel_client': safe(data.contact) + ' - ' + safe(data.contactTel)
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
  // Trouver la table de log (première table avec "date" dans l'en-tête)
  var tables = body.getTables();
  var logTable = null;
  for (var t = 0; t < tables.length; t++) {
    if (tables[t].getRow(0).getText().toLowerCase().indexOf('date') !== -1) {
      logTable = tables[t];
      break;
    }
  }

  // ---- INSTRUCTIONS DETAILLEES ----
  var instRange = body.findText('\\{\\{instructions\\}\\}');
  if (instRange) {
    var element = instRange.getElement();
    var paragraph = element.getParent().asParagraph();
    var container = paragraph.getParent();
    var insertionIndex = container.getChildIndex(paragraph);

    var allPostes = safeArray(data.postes);

    // Traiter TOUS les postes (simples et détaillés)
    allPostes.forEach(function(p, posteIdx) {
      var titre = safe(p.titre, 'PRESTATION ' + (posteIdx + 1)).toUpperCase();
      var nbJoursVal = parseFloat(p.jours) || 1;
      var joursFmt = nbJoursVal === 0.5 ? '1/2 jour' : (nbJoursVal + (nbJoursVal > 1 ? ' jours' : ' jour'));

      if (p.mode === 'simple') {
        // ---- POSTE SIMPLE : titre + description ----
        var pTitleS = container.insertParagraph(insertionIndex++, titre);
        pTitleS.editAsText().setBold(true).setForegroundColor('#000000').setUnderline(true);

        if (safe(p.text)) {
          var pTextS = container.insertParagraph(insertionIndex++, safe(p.text));
          pTextS.editAsText().setBold(false).setUnderline(false);
        }

        // Ligne vide de séparation
        container.insertParagraph(insertionIndex++, '');

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

        // Ligne RDV
        var rdvText = 'RDV : ' + rdvLines.join(' / ') + ' (' + joursFmt + ')';
        var pRDV = container.insertParagraph(insertionIndex++, rdvText);
        pRDV.editAsText().setBold(true).setForegroundColor(CONFIG.COLOR_PELICHET).setUnderline(true);

        // Titre du poste
        var pTitle = container.insertParagraph(insertionIndex++, titre);
        pTitle.editAsText().setBold(true).setForegroundColor('#000000').setUnderline(true);

        // Tache / instructions
        if (safe(p.tache)) {
          var pTask = container.insertParagraph(insertionIndex++, safe(p.tache));
          pTask.editAsText().setBold(false).setUnderline(false);
        }

        // Personnel
        var personnelList = transformerPersonnel(safeArray(p.personnel));
        if (personnelList.length > 0) {
          var pPers = container.insertParagraph(insertionIndex++, 'Personnel : ' + personnelList.join(', '));
          pPers.editAsText().setItalic(true).setFontSize(9);
        }

        // Vehicules
        var vehiculesList = safeArray(p.vehicules);
        if (vehiculesList.length > 0) {
          var pVeh = container.insertParagraph(insertionIndex++, 'Vehicules : ' + vehiculesList.join(', '));
          pVeh.editAsText().setItalic(true).setFontSize(9);
        }

        // Engins / vehicules speciaux
        var enginsList = safeArray(p.engins);
        if (enginsList.length > 0) {
          var pEng = container.insertParagraph(insertionIndex++, 'Vehicules speciaux : ' + enginsList.join(', '));
          pEng.editAsText().setItalic(true).setFontSize(9);
        }

        // Materiel
        var materielList = safeArray(p.materiel);
        if (materielList.length > 0) {
          var pMat = container.insertParagraph(insertionIndex++, 'Materiel : ' + materielList.join(', '));
          pMat.editAsText().setItalic(true).setFontSize(9);
        }

        // Ligne vide de séparation
        container.insertParagraph(insertionIndex++, '');

        // ---- Tableau ressources : 1 ligne par jour ouvre ----
        if (logTable) {
          // Date de debut = premier RDV saisi
          var dateDebut = (rdvs[0].date instanceof Date) ? rdvs[0].date : new Date(rdvs[0].date);
          var heureRDV = safe(rdvs[0].heure, '8H00');

          // Generer la liste des jours ouvres (excl. weekends + feries suisses)
          var joursOuvres = getJoursOuvres(dateDebut, nbJoursVal);

          // Gerer le demi-jour : la derniere ligne est "1/2 jour" si nbJours a une partie decimale
          var hasDemiJour = (nbJoursVal % 1 !== 0);

          joursOuvres.forEach(function(jourDate, jourIdx) {
            var dateFmtJour = Utilities.formatDate(jourDate, CONFIG.TIMEZONE, 'dd.MM.yyyy');
            var isLastDay = (jourIdx === joursOuvres.length - 1);
            var dureeJour = (isLastDay && hasDemiJour) ? '1/2 jour' : '1 jour';

            var row = logTable.appendTableRow();
            row.appendTableCell(dateFmtJour);
            row.appendTableCell(heureRDV);
            row.appendTableCell(personnelList.join('\n'));
            row.appendTableCell(dureeJour);
            row.appendTableCell(vehiculesList.join('\n'));
            row.appendTableCell(enginsList.join('\n'));
          });
        }
      }
    });

    // Supprimer le paragraphe {{instructions}}
    paragraph.removeFromParent();
  }

  docResa.saveAndClose();
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
