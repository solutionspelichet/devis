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
  TEMPLATE_RESA_ID: '11eE4_rUCztJvD6alGVav1Y1_yz0fyudj6i336rYQwkc',
  FOLDER_ID: '1DKVN4T2gKhPf26qOpT0sGiXrWHzSUBDk',
  SPREADSHEET_ID: '1AUfeykbUZ07SG-WkqVuxWJY6plH43EZp0tJgrp_gwYM',
  COLOR_PELICHET: '#D32F2F',
  TVA_RATE: 0.081,
  RPLP_RATE: 0.005,
  TIMEZONE: 'GMT+1',
  LOCK_TIMEOUT_MS: 30000,
  SHEET_NAME: 'Suivi_Devis'
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
 * Retourne le prochain jour ouvré après une date donnée
 */
function getNextWorkingDay(date) {
  const next = new Date(date.getTime());
  next.setDate(next.getDate() + 1);
  while (next.getDay() === 0 || next.getDay() === 6) {
    next.setDate(next.getDate() + 1);
  }
  return next;
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

    if (action === 'rechercher') {
      const ref = e.parameter.ref || '';
      const result = rechercherDossier(ref);
      if (result) {
        return jsonResponse({ status: 'success', data: result });
      }
      return jsonResponse({ status: 'not_found' });
    }

    if (action === 'lister') {
      const dossiers = listerDossiers();
      return jsonResponse({ status: 'success', data: dossiers });
    }

    return jsonResponse({ status: 'error', message: 'Action inconnue' });
  } catch (err) {
    Logger.log('ERREUR doGet: ' + err);
    return jsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Trouve la feuille de suivi (par nom configuré, sinon première feuille)
 */
function getSheet() {
  const ss = getSs();
  if (!ss) return null;
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    const sheets = ss.getSheets();
    if (sheets.length === 0) return null;
    sheet = sheets[0];
    Logger.log('Feuille "' + CONFIG.SHEET_NAME + '" introuvable, utilisation de "' + sheet.getName() + '"');
  }
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
function listerDossiers() {
  const sheet = getSheet();
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

function rechercherDossier(ref) {
  const sheet = getSheet();
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
// TRAITEMENT PRINCIPAL (POST)
// ============================================
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(CONFIG.LOCK_TIMEOUT_MS)) {
      return jsonResponse({ status: 'error', message: 'Serveur occupé, réessayez dans quelques secondes.' });
    }

    const data = JSON.parse(e.postData.contents);

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

    const replacements = {
      '{{nom_societe}}': estPelichet ? 'PELICHET NLC SA' : safe(data.nomPrestataire, 'Société'),
      '{{ref}}': ref,
      '{{date}}': dateJour(),
      '{{salutation}}': safe(data.genre, 'Monsieur'),
      '{{vendeur}}': safe(data.vendeur, 'ARNAUD GUEDOU'),
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
    docDevis.saveAndClose();

    // 2. FICHE RESA (Pelichet uniquement)
    if (estPelichet) {
      try {
        genererFicheResa(data, folder, ref, client);
        Logger.log('Fiche RESA générée: ' + ref);
      } catch (err) {
        Logger.log('ERREUR RESA (non bloquante): ' + err);
      }
    }

    // 3. PDF + ARCHIVAGE
    const pdfBlob = copyDevis.getAs('application/pdf');
    const pdfFile = folder.createFile(pdfBlob).setName(fileNameBase + '.pdf');

    archiver(ref, estPelichet, client, data, montantHT, pdfFile);

    Logger.log('Devis terminé: ' + ref + ' -> ' + pdfFile.getUrl());

    return jsonResponse({
      status: 'success',
      pdfUrl: pdfFile.getUrl(),
      docUrl: copyDevis.getUrl()
    });

  } catch (error) {
    Logger.log('ERREUR CRITIQUE doPost: ' + error + '\n' + error.stack);
    return jsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ============================================
// ARCHIVAGE SPREADSHEET
// ============================================
function archiver(ref, estPelichet, client, data, montantHT, pdfFile) {
  try {
    const ss = getSs();
    if (!ss) return;
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      // Chercher la première feuille disponible
      var sheets = ss.getSheets();
      sheet = sheets.length > 0 ? sheets[0] : ss.insertSheet(CONFIG.SHEET_NAME);
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
    'vendeur': safe(data.vendeur, 'ARNAUD GUEDOU').toUpperCase(),
    'client': client,
    'adresse_depart': safe(data.adresseDepart),
    'adresse_arrivee': safe(data.adresseArrivee),
    'volume': volumeEstime + ' M3',
    'tel_client': safe(data.contact) + ' - ' + safe(data.contactTel),
    'vendeur_tel': safe(data.vendeurTel, '+41 79 688 27 35')
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

        // ---- Ligne dans le tableau ressources (1 ligne par poste, pas par jour) ----
        if (logTable) {
          var row = logTable.appendTableRow();
          row.appendTableCell(rdvLines.join('\n'));
          row.appendTableCell(safe(rdvs[0].heure, '8H00'));
          row.appendTableCell(personnelList.join('\n'));
          row.appendTableCell(joursFmt);
          row.appendTableCell(vehiculesList.join('\n'));
          row.appendTableCell(enginsList.join('\n'));
        }
      }
    });

    // Supprimer le paragraphe {{instructions}}
    paragraph.removeFromParent();
  }

  docResa.saveAndClose();
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
