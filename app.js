/**
 * PELICHET NLC SA — Application Devis & Logistique v2.0
 */

// ============================================
// CONFIGURATION
// ============================================
const CONFIG = {
  SCRIPT_URL: 'https://script.google.com/macros/s/AKfycbxYONsnng1aenrN00UsInpYTYN-oy1P5SVS9unxkahyzLCNPKWT7IH-4IR33R-ucQ8Kzg/exec',
  TVA_RATE: 0.081,
  RPLP_RATE: 0.005,
  LISTS: {
    personnel: ['Manutentionnaire', 'Manutentionnaire du lourd', 'Chauffeur', "Chef d'équipe", 'CHAUF-LIVREUR'],
    vehicules: ['VL', '1F', 'PL', 'Semi', 'Box IT'],
    engins: ['Chariot élévateur', 'Grue mobile', 'Monte-meuble'],
    materiel: ['Chariots', 'Rouleaux bulle', 'Adhésif', 'Transpalette']
  }
};

// ============================================
// CUSTOM ITEMS (persistance localStorage)
// ============================================
const CustomItems = {
  _cache: null,
  _loaded: false,

  /** Charge les items depuis le serveur (avec cache local) */
  load() {
    // Retourne le cache synchrone (chargé au démarrage)
    if (this._cache) return this._cache;
    // Fallback localStorage en attendant le serveur
    try {
      const raw = localStorage.getItem('pelichet_custom_items');
      if (raw) return JSON.parse(raw);
    } catch (e) { /* silencieux */ }
    return { personnel: [], vehicules: [], engins: [], materiel: [] };
  },

  /** Charge depuis le serveur et met à jour le cache */
  async fetchFromServer() {
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=custom_get`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && result.data) {
        this._cache = result.data;
        this._loaded = true;
        // Sauvegarder en localStorage comme fallback
        localStorage.setItem('pelichet_custom_items', JSON.stringify(this._cache));
      }
    } catch (e) {
      console.log('CustomItems fetch failed, using localStorage fallback');
    }
  },

  /** Ajoute un item sur le serveur + cache local */
  async add(type, value) {
    if (!value || !value.trim()) return false;
    const val = value.trim();
    // Vérifier qu'il n'est pas déjà dans la liste par défaut
    if (CONFIG.LISTS[type] && CONFIG.LISTS[type].includes(val)) return false;
    // Vérifier doublon dans le cache
    const data = this.load();
    if (data[type] && data[type].includes(val)) return false;

    // Ajouter au cache local immédiatement
    if (!data[type]) data[type] = [];
    data[type].push(val);
    this._cache = data;
    localStorage.setItem('pelichet_custom_items', JSON.stringify(data));

    // Envoyer au serveur en arrière-plan
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=custom_add&type=${encodeURIComponent(type)}&value=${encodeURIComponent(val)}`;
      fetch(url).catch(() => {}); // fire & forget
    } catch (e) { /* silencieux */ }

    return true;
  },

  /** Retourne la liste complète (défaut + custom) pour un type */
  getFullList(type) {
    const defaults = CONFIG.LISTS[type] || [];
    const customs = this.load()[type] || [];
    return [...defaults, ...customs];
  }
};

// ============================================
// TARIF MANAGER (tarifs + marges depuis serveur)
// ============================================
const TarifManager = {
  _data: null, // { items: { personnel: [{item, cout, unite},...] }, marges: { personnel: 35, ... } }
  _loaded: false,

  async fetchFromServer() {
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=tarifs_get`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && result.data) {
        this._data = result.data;
        this._loaded = true;
        localStorage.setItem('pelichet_tarifs', JSON.stringify(this._data));
      }
    } catch (e) {
      console.log('TarifManager fetch failed, using localStorage');
    }
    if (!this._data) {
      try {
        const raw = localStorage.getItem('pelichet_tarifs');
        if (raw) this._data = JSON.parse(raw);
      } catch (e) { /* silencieux */ }
    }
    if (!this._data) this._data = { items: {}, marges: {} };
  },

  async saveToServer() {
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=tarifs_save&data=${encodeURIComponent(JSON.stringify(this._data))}`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success') {
        localStorage.setItem('pelichet_tarifs', JSON.stringify(this._data));
        return true;
      }
    } catch (e) { /* silencieux */ }
    return false;
  },

  /** Retourne le coût unitaire d'un item */
  getCout(type, itemName) {
    if (!this._data || !this._data.items || !this._data.items[type]) return 0;
    const found = this._data.items[type].find(t => t.item === itemName);
    return found ? (parseFloat(found.cout) || 0) : 0;
  },

  /** Retourne l'unité d'un item ('jour', 'piece', 'forfait') */
  getUnite(type, itemName) {
    if (!this._data || !this._data.items || !this._data.items[type]) return 'jour';
    const found = this._data.items[type].find(t => t.item === itemName);
    return found ? (found.unite || 'jour') : 'jour';
  },

  /** Retourne la marge % d'une catégorie */
  getMarge(type) {
    if (!this._data || !this._data.marges) return 0;
    return parseFloat(this._data.marges[type]) || 0;
  },

  /** Calcule le prix d'un poste détail à partir de ses ressources */
  calculerPrixPoste(card) {
    const nbJours = parseFloat(card.querySelector('[name="nbJours"]')?.value) || 1;
    const types = ['engins', 'personnel', 'vehicules', 'materiel'];
    const sections = card.querySelectorAll('.detail-section .rows');
    let totalCout = 0;
    let totalPrix = 0;
    const details = [];

    types.forEach((type, idx) => {
      const rows = sections[idx]?.querySelectorAll('.row-item') || [];
      let coutCategorie = 0;

      rows.forEach(row => {
        const selectVal = row.querySelector('select')?.value || '';
        const customVal = row.querySelector('input[name="customValue"]')?.value || '';
        const label = (selectVal === '__autre__') ? customVal : selectVal;
        if (!label) return;

        const qty = parseFloat(row.querySelector('input[name="qty"]')?.value) || 1;
        const coutUnit = this.getCout(type, label);
        const unite = this.getUnite(type, label);

        // Jour = multiplié par nbJours, pièce/forfait = quantité seule
        const coutItem = unite === 'jour' ? (qty * coutUnit * nbJours) : (qty * coutUnit);
        coutCategorie += coutItem;
      });

      const marge = this.getMarge(type);
      const prixCategorie = coutCategorie * (1 + marge / 100);
      totalCout += coutCategorie;
      totalPrix += prixCategorie;

      if (coutCategorie > 0) {
        details.push({ type, cout: coutCategorie, marge, prix: prixCategorie });
      }
    });

    return { cout: Math.round(totalCout * 100) / 100, prix: Math.round(totalPrix * 100) / 100, details };
  },

  /** Met à jour les données en mémoire */
  setData(data) {
    this._data = data;
  },

  getData() {
    return this._data || { items: {}, marges: {} };
  }
};

// ============================================
// SETTINGS PANEL (édition des tarifs)
// ============================================
const SettingsPanel = {
  _panel: null,
  _open: false,

  init() {
    this._panel = document.getElementById('settingsPanel');
    document.getElementById('settingsBtn')?.addEventListener('click', () => this.toggle());
    document.getElementById('settingsClose')?.addEventListener('click', () => this.close());
    document.getElementById('settingsSave')?.addEventListener('click', () => this.save());
    document.getElementById('settingsAddRow')?.addEventListener('click', () => this.addRow());
  },

  toggle() {
    this._open ? this.close() : this.open();
  },

  open() {
    this._open = true;
    this._panel.classList.remove('hidden');
    this.render();
  },

  close() {
    this._open = false;
    this._panel.classList.add('hidden');
  },

  render() {
    const data = TarifManager.getData();
    const marges = data.marges || {};
    const items = data.items || {};

    // Marges
    let margesHTML = '';
    ['personnel', 'vehicules', 'engins', 'materiel'].forEach(cat => {
      margesHTML += `
        <div class="settings-marge-row">
          <span class="settings-marge-label">${cat}</span>
          <input type="number" data-marge="${cat}" value="${marges[cat] || 0}" min="0" max="200" step="1">
          <span class="settings-marge-unit">%</span>
        </div>`;
    });
    document.getElementById('settingsMarges').innerHTML = margesHTML;

    // Tarifs
    let tarifsHTML = '';
    const allTypes = ['personnel', 'vehicules', 'engins', 'materiel'];
    allTypes.forEach(type => {
      const list = items[type] || [];
      list.forEach((t, idx) => {
        tarifsHTML += `
          <tr data-type="${type}" data-idx="${idx}">
            <td><select class="tarif-type">${allTypes.map(tt => `<option value="${tt}" ${tt === type ? 'selected' : ''}>${tt}</option>`).join('')}</select></td>
            <td><input type="text" class="tarif-item" value="${t.item || ''}"></td>
            <td><input type="number" class="tarif-cout" value="${t.cout || 0}" min="0" step="0.5"></td>
            <td><select class="tarif-unite"><option value="jour" ${t.unite === 'jour' ? 'selected' : ''}>jour</option><option value="piece" ${t.unite === 'piece' ? 'selected' : ''}>piece</option><option value="forfait" ${t.unite === 'forfait' ? 'selected' : ''}>forfait</option></select></td>
            <td><button type="button" class="tarif-remove" onclick="this.closest('tr').remove()">&times;</button></td>
          </tr>`;
      });
    });
    document.getElementById('settingsTarifsBody').innerHTML = tarifsHTML;
  },

  addRow() {
    const tbody = document.getElementById('settingsTarifsBody');
    const allTypes = ['personnel', 'vehicules', 'engins', 'materiel'];
    tbody.insertAdjacentHTML('beforeend', `
      <tr data-type="" data-idx="new">
        <td><select class="tarif-type">${allTypes.map(tt => `<option value="${tt}">${tt}</option>`).join('')}</select></td>
        <td><input type="text" class="tarif-item" value="" placeholder="Nom item"></td>
        <td><input type="number" class="tarif-cout" value="0" min="0" step="0.5"></td>
        <td><select class="tarif-unite"><option value="jour">jour</option><option value="piece">piece</option><option value="forfait">forfait</option></select></td>
        <td><button type="button" class="tarif-remove" onclick="this.closest('tr').remove()">&times;</button></td>
      </tr>`);
  },

  async save() {
    const data = { items: {}, marges: {} };

    // Lire les marges
    document.querySelectorAll('[data-marge]').forEach(input => {
      data.marges[input.dataset.marge] = parseFloat(input.value) || 0;
    });

    // Lire les tarifs
    document.querySelectorAll('#settingsTarifsBody tr').forEach(tr => {
      const type = tr.querySelector('.tarif-type')?.value;
      const item = tr.querySelector('.tarif-item')?.value?.trim();
      const cout = parseFloat(tr.querySelector('.tarif-cout')?.value) || 0;
      const unite = tr.querySelector('.tarif-unite')?.value || 'jour';
      if (!type || !item) return;
      if (!data.items[type]) data.items[type] = [];
      data.items[type].push({ item, cout, unite });
    });

    TarifManager.setData(data);
    Toast.info('Sauvegarde en cours...');
    const ok = await TarifManager.saveToServer();
    if (ok) {
      Toast.success('Tarifs sauvegardés !');
    } else {
      Toast.error('Erreur de sauvegarde');
    }
  }
};

// ============================================
// TOAST NOTIFICATIONS
// ============================================
const Toast = {
  _container: null,

  _getContainer() {
    if (!this._container) {
      this._container = document.createElement('div');
      this._container.className = 'toast-container';
      document.body.appendChild(this._container);
    }
    return this._container;
  },

  show(message, type = 'info', duration = 4000) {
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    this._getContainer().appendChild(toast);
    setTimeout(() => toast.remove(), duration);
  },

  success(msg) { this.show(msg, 'success'); },
  error(msg) { this.show(msg, 'error', 6000); },
  info(msg) { this.show(msg, 'info'); },
  warning(msg) { this.show(msg, 'warning'); }
};

// ============================================
// FORM VALIDATOR
// ============================================
const Validator = {
  rules: {
    ref: { required: true, message: 'Référence dossier requise' },
    client: { required: true, message: 'Nom du client requis' },
    adresseClient: { required: true, message: 'Adresse de facturation requise' },
    adresseDepart: { required: true, message: 'Adresse de départ requise' },
    adresseArrivee: { required: true, message: "Adresse d'arrivée requise" },
    montantHT: { required: true, min: 0.01, message: 'Montant HT requis (> 0)' }
  },

  validate(form) {
    this.clearErrors(form);
    let valid = true;

    for (const [name, rule] of Object.entries(this.rules)) {
      const field = form.querySelector(`[name="${name}"]`);
      if (!field) continue;

      const value = field.value.trim();
      let hasError = false;

      if (rule.required && !value) hasError = true;
      if (rule.min && (parseFloat(value) < rule.min || isNaN(parseFloat(value)))) hasError = true;

      if (hasError) {
        valid = false;
        field.classList.add('error');
        const errEl = field.parentElement.querySelector('.error-msg');
        if (errEl) { errEl.textContent = rule.message; errEl.classList.add('visible'); }
      }
    }

    const postes = document.querySelectorAll('.poste-card');
    if (postes.length === 0) {
      valid = false;
      Toast.warning('Ajoutez au moins un poste de prestation.');
    }

    return valid;
  },

  clearErrors(form) {
    form.querySelectorAll('.error').forEach(el => el.classList.remove('error'));
    form.querySelectorAll('.error-msg').forEach(el => el.classList.remove('visible'));
  }
};

// ============================================
// PRICE CALCULATOR
// ============================================
const PriceCalc = {
  _htInput: null,
  _breakdownEl: null,

  init() {
    this._htInput = document.querySelector('[name="montantHT"]');
    this._breakdownEl = document.getElementById('totalBreakdown');
    if (this._htInput) {
      this._htInput.addEventListener('input', () => this.updateBreakdown());
    }
    this.updateBreakdown();
  },

  sumPostes() {
    let total = 0;
    document.querySelectorAll('[name="postePrix"]').forEach(input => {
      const val = parseFloat(input.value);
      if (!isNaN(val)) total += val;
    });
    return total;
  },

  applyAutoTotal() {
    const sum = this.sumPostes();
    if (this._htInput) {
      this._htInput.value = sum.toFixed(2);
      this.updateBreakdown();
    }
  },

  updateBreakdown() {
    if (!this._breakdownEl || !this._htInput) return;
    const ht = parseFloat(this._htInput.value) || 0;
    const tva = ht * CONFIG.TVA_RATE;
    const rplp = ht * CONFIG.RPLP_RATE;
    const ttc = ht + tva + rplp;
    const fmt = (v) => v.toLocaleString('fr-CH', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' CHF';

    this._breakdownEl.innerHTML = `
      <div class="line"><span>Total HT</span><span>${fmt(ht)}</span></div>
      <div class="line"><span>TVA (8.1%)</span><span>${fmt(tva)}</span></div>
      <div class="line"><span>RPLP (0.5%)</span><span>${fmt(rplp)}</span></div>
      <div class="line total"><span>Total TTC</span><span>${fmt(ttc)}</span></div>
    `;
  }
};

// ============================================
// POSTE MANAGER
// ============================================
const PosteManager = {
  _container: null,

  init() {
    this._container = document.getElementById('prestationsContainer');
    this.addPoste('simple');
  },

  _createOptionsHTML(type) {
    const fullList = CustomItems.getFullList(type);
    const defaults = CONFIG.LISTS[type] || [];
    const customs = CustomItems.load()[type] || [];

    let html = '<option value="">-- Type --</option>';
    // Options par défaut
    html += defaults.map(i => `<option value="${i}">${i}</option>`).join('');
    // Séparateur + options personnalisées
    if (customs.length > 0) {
      html += '<option disabled>────────────</option>';
      html += customs.map(i => `<option value="${i}">${i} ★</option>`).join('');
    }
    // Option "Autre"
    html += '<option value="__autre__">+ Autre (personnalisé)</option>';
    return html;
  },

  /** Rafraîchit toutes les listes déroulantes d'un type donné */
  _refreshSelects(type) {
    const newHTML = this._createOptionsHTML(type);
    document.querySelectorAll(`.row-item[data-type="${type}"] select`).forEach(sel => {
      const currentVal = sel.value;
      sel.innerHTML = newHTML;
      if (currentVal && currentVal !== '__autre__') sel.value = currentVal;
    });
  },

  _createRow(type) {
    const isEngin = type === 'engins';
    const div = document.createElement('div');
    div.className = 'row-item';
    div.dataset.type = type;
    div.innerHTML = `
      <select style="flex:1">${this._createOptionsHTML(type)}</select>
      <input type="text" name="customValue" placeholder="Précisez..." class="custom-input hidden" style="flex:1">
      ${isEngin ? `
        <div style="position:relative;display:flex;align-items:center">
          <input type="text" name="ton" placeholder="Ton" style="width:2.5rem;text-align:center">
          <span style="font-size:8px;font-weight:700;color:var(--slate-400);margin-left:2px">T</span>
        </div>
        <div style="display:flex;align-items:center;margin-left:0.5rem">
          <input type="number" name="qty" placeholder="Qté" value="1" style="width:2.5rem;text-align:center">
          <span style="font-size:8px;font-weight:700;color:var(--slate-400);margin-left:2px">x</span>
        </div>
      ` : `<input type="number" name="qty" value="1" style="width:3rem;text-align:center">`}
      <button type="button" class="row-remove" title="Supprimer">&times;</button>
    `;
    const select = div.querySelector('select');
    const customInput = div.querySelector('.custom-input');
    const self = this;

    // Toggle champ personnalisé quand "Autre" est sélectionné
    select.addEventListener('change', () => {
      if (select.value === '__autre__') {
        customInput.classList.remove('hidden');
        customInput.focus();
      } else {
        customInput.classList.add('hidden');
        customInput.value = '';
      }
    });

    // Enregistrer la valeur personnalisée quand on quitte le champ
    customInput.addEventListener('blur', async () => {
      const val = customInput.value.trim();
      if (!val) return;
      const added = await CustomItems.add(type, val);
      if (added) {
        Toast.info('"' + val + '" ajouté à la liste ' + type);
      }
      // Rafraîchir tous les selects de ce type + sélectionner la nouvelle valeur
      self._refreshSelects(type);
      select.value = val;
      customInput.classList.add('hidden');
      customInput.value = '';
    });

    // Aussi enregistrer sur Entrée
    customInput.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        e.preventDefault();
        customInput.blur();
      }
    });

    div.querySelector('.row-remove').addEventListener('click', () => div.remove());
    return div;
  },

  _createRDVRow() {
    const div = document.createElement('div');
    div.className = 'rdv-row';
    div.innerHTML = `
      <input type="date" name="posteDate" style="flex:1">
      <input type="text" name="heureRDV" value="8H00" style="width:4rem;font-weight:700">
      <button type="button" class="row-remove" title="Supprimer">&times;</button>
    `;
    div.querySelector('.row-remove').addEventListener('click', () => div.remove());
    return div;
  },

  addPoste(mode = 'simple') {
    const div = document.createElement('div');
    div.className = `poste-card poste-${mode}`;
    div.dataset.mode = mode;

    div.innerHTML = `
      <span class="poste-number"></span>
      <button type="button" class="poste-remove" title="Supprimer le poste">&times;</button>
      <div class="grid-2 mb-3">
        <div class="col-span-2" style="display:grid;grid-template-columns:1fr auto;gap:0.5rem">
          <input type="text" name="posteTitre" placeholder="DESIGNATION" class="font-bold uppercase">
          <input type="number" name="postePrix" placeholder="Prix HT" style="width:7rem;text-align:right" class="font-bold">
        </div>
      </div>
      ${mode === 'simple' ? `
        <textarea name="simpleText" rows="3" placeholder="Description libre..." class="text-sm"></textarea>
      ` : `
        <div class="space-y-4">
          <div class="rdv-section">
            <label>RDV & Durée :</label>
            <div class="rdv-container"></div>
            <button type="button" class="add-row-btn rdv mt-1">+ Ajouter jour</button>
            <div class="jours-row">
              <label>Nombre de jours :</label>
              <input type="number" name="nbJours" step="0.5" value="1" min="0.5">
              <button type="button" class="half-btn" data-half>0.5 J</button>
              <button type="button" class="half-btn" data-one>1 J</button>
              <button type="button" class="half-btn" data-two>2 J</button>
            </div>
          </div>
          <div class="detail-section engins">
            <label>Véhicules Spéciaux :</label>
            <div class="rows"></div>
            <button type="button" class="add-row-btn engins">+ Ajouter</button>
          </div>
          <div class="detail-section personnel">
            <label>Personnel :</label>
            <div class="rows"></div>
            <button type="button" class="add-row-btn personnel">+ Ajouter</button>
          </div>
          <div class="detail-section vehicules">
            <label>Véhicules :</label>
            <div class="rows"></div>
            <button type="button" class="add-row-btn vehicules">+ Ajouter</button>
          </div>
          <div class="detail-section materiel">
            <label>Matériel :</label>
            <div class="rows"></div>
            <button type="button" class="add-row-btn materiel">+ Ajouter</button>
          </div>
          <textarea name="tache" rows="2" placeholder="Instructions..." class="text-sm"></textarea>
          <div class="prix-auto-info hidden">
            <div class="prix-auto-detail"></div>
            <button type="button" class="prix-auto-apply">Appliquer le prix calculé</button>
          </div>
        </div>
      `}
    `;

    // Remove button
    div.querySelector('.poste-remove').addEventListener('click', () => {
      div.remove();
      this.renumber();
      PriceCalc.updateBreakdown();
    });

    // Prix change -> update breakdown + indicateur écart
    const prixInput = div.querySelector('[name="postePrix"]');
    if (prixInput) {
      prixInput.addEventListener('input', () => {
        prixInput.dataset.manual = 'true';
        PriceCalc.updateBreakdown();
        this._updatePrixIndicator(div);
      });
    }

    // Detail mode event bindings
    if (mode === 'detail') {
      const rdvContainer = div.querySelector('.rdv-container');
      rdvContainer.appendChild(this._createRDVRow());

      div.querySelector('.add-row-btn.rdv').addEventListener('click', () => {
        rdvContainer.appendChild(this._createRDVRow());
      });

      const nbJoursInput = div.querySelector('[name="nbJours"]');
      const recalc = () => this._recalculerPoste(div);
      div.querySelector('[data-half]').addEventListener('click', () => { nbJoursInput.value = '0.5'; recalc(); });
      div.querySelector('[data-one]').addEventListener('click', () => { nbJoursInput.value = '1'; recalc(); });
      div.querySelector('[data-two]').addEventListener('click', () => { nbJoursInput.value = '2'; recalc(); });
      nbJoursInput.addEventListener('input', recalc);

      // Bouton "Appliquer prix calculé"
      div.querySelector('.prix-auto-apply')?.addEventListener('click', () => {
        const calc = TarifManager.calculerPrixPoste(div);
        const pi = div.querySelector('[name="postePrix"]');
        if (pi) { pi.value = calc.prix.toFixed(2); pi.dataset.manual = ''; PriceCalc.updateBreakdown(); }
        this._updatePrixIndicator(div);
      });

      ['engins', 'personnel', 'vehicules', 'materiel'].forEach(type => {
        const section = div.querySelector(`.detail-section.${type}`);
        section.querySelector('.add-row-btn').addEventListener('click', () => {
          const row = this._createRow(type);
          section.querySelector('.rows').appendChild(row);
          // Recalculer quand on change select ou qty dans la nouvelle ligne
          row.querySelector('select')?.addEventListener('change', recalc);
          row.querySelector('input[name="qty"]')?.addEventListener('input', recalc);
        });
      });
    }

    this._container.appendChild(div);
    this.renumber();
  },

  /** Recalcule le prix d'un poste détail et met à jour l'affichage */
  _recalculerPoste(card) {
    const calc = TarifManager.calculerPrixPoste(card);
    const prixInput = card.querySelector('[name="postePrix"]');
    if (!prixInput) return;

    // Si pas de modification manuelle, pré-remplir
    if (!prixInput.dataset.manual && calc.prix > 0) {
      prixInput.value = calc.prix.toFixed(2);
      PriceCalc.updateBreakdown();
    }

    this._updatePrixIndicator(card);
  },

  /** Met à jour l'indicateur de prix calculé sous le champ prix */
  _updatePrixIndicator(card) {
    const infoDiv = card.querySelector('.prix-auto-info');
    const detailDiv = card.querySelector('.prix-auto-detail');
    if (!infoDiv || !detailDiv) return;

    const calc = TarifManager.calculerPrixPoste(card);
    if (calc.prix === 0 && calc.cout === 0) {
      infoDiv.classList.add('hidden');
      return;
    }

    infoDiv.classList.remove('hidden');
    const fmtCHF = v => v.toLocaleString('fr-CH', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
    const prixManuel = parseFloat(card.querySelector('[name="postePrix"]')?.value) || 0;
    const ecart = prixManuel - calc.prix;
    const ecartPct = calc.prix > 0 ? ((ecart / calc.prix) * 100).toFixed(1) : 0;

    let html = `<div class="prix-auto-ligne"><span>Cout de revient</span><span>${fmtCHF(calc.cout)} CHF</span></div>`;
    calc.details.forEach(d => {
      html += `<div class="prix-auto-ligne sub"><span>${d.type} (+${d.marge}%)</span><span>${fmtCHF(d.prix)} CHF</span></div>`;
    });
    html += `<div class="prix-auto-ligne total"><span>Prix calculé</span><span>${fmtCHF(calc.prix)} CHF</span></div>`;

    if (prixManuel > 0 && Math.abs(ecart) > 1) {
      const cls = ecart > 0 ? 'ecart-positif' : 'ecart-negatif';
      html += `<div class="prix-auto-ligne ${cls}"><span>Ecart</span><span>${ecart > 0 ? '+' : ''}${fmtCHF(ecart)} CHF (${ecartPct}%)</span></div>`;
    }

    detailDiv.innerHTML = html;
  },

  renumber() {
    this._container.querySelectorAll('.poste-card').forEach((card, idx) => {
      const num = card.querySelector('.poste-number');
      if (num) num.textContent = idx + 1;
    });
  },

  collectAll() {
    return Array.from(this._container.querySelectorAll('.poste-card')).map(card => {
      const mode = card.dataset.mode;
      const base = {
        mode,
        titre: card.querySelector('[name="posteTitre"]')?.value || '',
        prix: card.querySelector('[name="postePrix"]')?.value || ''
      };

      if (mode === 'detail') {
        base.jours = card.querySelector('[name="nbJours"]')?.value || '1';
        base.tache = card.querySelector('[name="tache"]')?.value || '';
        base.rdvs = Array.from(card.querySelectorAll('.rdv-row')).map(r => ({
          date: r.querySelector('[name="posteDate"]')?.value || '',
          heure: r.querySelector('[name="heureRDV"]')?.value || '8H00'
        }));

        const sections = card.querySelectorAll('.detail-section .rows');
        base.engins = this._collectRows(sections[0], true);
        base.personnel = this._collectRows(sections[1]);
        base.vehicules = this._collectRows(sections[2]);
        base.materiel = this._collectRows(sections[3]);
      } else {
        base.text = card.querySelector('[name="simpleText"]')?.value || '';
      }
      return base;
    });
  },

  _collectRows(container, hasEnginFields = false) {
    if (!container) return [];
    return Array.from(container.querySelectorAll('.row-item')).map(r => {
      const selectVal = r.querySelector('select')?.value || '';
      const customVal = r.querySelector('input[name="customValue"]')?.value || '';
      // Si "Autre" est sélectionné, utiliser la valeur personnalisée
      const label = (selectVal === '__autre__') ? customVal : selectVal;
      if (!label.trim()) return '';
      const qty = r.querySelector('input[name="qty"]')?.value || '1';
      if (hasEnginFields) {
        const ton = r.querySelector('input[name="ton"]')?.value || '';
        return `${qty}x ${label} (${ton}T)`;
      }
      return `${qty}x ${label}`;
    }).filter(s => s);
  },

  loadPostes(postes) {
    this._container.innerHTML = '';
    (postes || []).forEach(p => {
      const mode = p.mode || 'simple';
      this.addPoste(mode);
      const card = this._container.lastElementChild;
      card.querySelector('[name="posteTitre"]').value = p.titre || '';
      card.querySelector('[name="postePrix"]').value = p.prix || '';

      if (mode === 'simple') {
        const ta = card.querySelector('[name="simpleText"]');
        if (ta) ta.value = p.text || '';

      } else {
        // ---- Nombre de jours ----
        const nbJoursInput = card.querySelector('[name="nbJours"]');
        if (nbJoursInput) nbJoursInput.value = p.jours || '1';

        // ---- Instructions / tâche ----
        const tacheInput = card.querySelector('[name="tache"]');
        if (tacheInput) tacheInput.value = p.tache || '';

        // ---- RDV (dates et heures) ----
        const rdvContainer = card.querySelector('.rdv-container');
        const rdvs = (p.rdvs || []).filter(r => r.date);
        if (rdvs.length > 0 && rdvContainer) {
          // Supprimer le RDV par défaut créé par addPoste
          rdvContainer.innerHTML = '';
          rdvs.forEach(r => {
            const rdvRow = this._createRDVRow();
            const dateInput = rdvRow.querySelector('[name="posteDate"]');
            const heureInput = rdvRow.querySelector('[name="heureRDV"]');
            if (dateInput) dateInput.value = r.date || '';
            if (heureInput) heureInput.value = r.heure || '8H00';
            rdvContainer.appendChild(rdvRow);
          });
        }

        // ---- Ressources (personnel, véhicules, engins, matériel) ----
        const sections = card.querySelectorAll('.detail-section .rows');
        const resourceTypes = ['engins', 'personnel', 'vehicules', 'materiel'];

        resourceTypes.forEach((type, idx) => {
          const items = p[type] || [];
          const rowsContainer = sections[idx];
          if (!rowsContainer || items.length === 0) return;

          items.forEach(itemStr => {
            // Parser le format "2x Manutentionnaire" ou "1x Grue mobile (50T)"
            const match = String(itemStr).match(/^(\d+)x\s+(.+?)(?:\s+\((\d+)T\))?$/);
            if (!match) return;

            const qty = match[1];
            const label = match[2].trim();
            const ton = match[3] || '';

            const row = this._createRow(type);
            const select = row.querySelector('select');
            const qtyInput = row.querySelector('input[name="qty"]');

            // Essayer de sélectionner dans la liste (défaut + custom)
            if (select) {
              const option = Array.from(select.options).find(o => o.value === label);
              if (option) {
                select.value = label;
              } else {
                // Valeur inconnue : l'enregistrer en custom puis sélectionner
                CustomItems.add(type, label); // async mais pas besoin d'attendre
                // Recréer les options avec la nouvelle valeur
                select.innerHTML = this._createOptionsHTML(type);
                select.value = label;
              }
            }

            if (qtyInput) qtyInput.value = qty;

            // Tonnage pour les engins
            if (ton && type === 'engins') {
              const tonInput = row.querySelector('input[name="ton"]');
              if (tonInput) tonInput.value = ton;
            }

            rowsContainer.appendChild(row);
          });
        });
      }
    });
    PriceCalc.updateBreakdown();
  }
};

// ============================================
// CONFIRMATION MODAL
// ============================================
const Modal = {
  show(data, onConfirm) {
    const overlay = document.createElement('div');
    overlay.className = 'modal-overlay';

    const postes = data.postes || [];
    const postesRecap = postes.map((p, i) =>
      `<div style="font-size:12px;padding:0.25rem 0">${i+1}. ${p.titre || 'Sans titre'} — ${p.prix ? p.prix + ' CHF' : 'Inclus'}</div>`
    ).join('');

    const ht = parseFloat(data.montantHT) || 0;
    const tva = ht * CONFIG.TVA_RATE;
    const rplp = ht * CONFIG.RPLP_RATE;
    const ttc = ht + tva + rplp;
    const fmt = (v) => v.toLocaleString('fr-CH', { minimumFractionDigits: 2 }) + ' CHF';

    overlay.innerHTML = `
      <div class="modal">
        <h3>Confirmer l'envoi du devis</h3>
        <div class="recap-item"><strong>Référence</strong>${data.ref}</div>
        <div class="recap-item"><strong>Client</strong>${data.client}</div>
        <div class="recap-item"><strong>Départ</strong>${data.adresseDepart}</div>
        <div class="recap-item"><strong>Arrivée</strong>${data.adresseArrivee}</div>
        <div class="recap-item"><strong>Postes (${postes.length})</strong>${postesRecap}</div>
        <div class="recap-item"><strong>Total TTC</strong>${fmt(ttc)}</div>
        <div class="modal-actions">
          <button type="button" class="btn" style="background:var(--slate-200);color:var(--slate-700)" id="modalCancel">Annuler</button>
          <button type="button" class="btn btn-primary" id="modalConfirm">Confirmer & Envoyer</button>
        </div>
      </div>
    `;

    document.body.appendChild(overlay);
    overlay.querySelector('#modalCancel').addEventListener('click', () => overlay.remove());
    overlay.querySelector('#modalConfirm').addEventListener('click', () => {
      overlay.remove();
      onConfirm();
    });
    overlay.addEventListener('click', (e) => { if (e.target === overlay) overlay.remove(); });
  }
};

// ============================================
// DOSSIER LIST (liste tous les dossiers)
// ============================================
const DossierList = {
  _panel: null,
  _body: null,
  _filter: null,
  _dossiers: [],
  _loaded: false,

  init() {
    this._panel = document.getElementById('dossierListPanel');
    this._body = document.getElementById('dossierListBody');
    this._filter = document.getElementById('dossierFilter');

    document.getElementById('listBtn')?.addEventListener('click', () => this.toggle());
    document.getElementById('closeDossierList')?.addEventListener('click', () => this.close());
    this._filter?.addEventListener('input', () => this._render());
  },

  toggle() {
    if (this._panel.classList.contains('hidden')) {
      this.open();
    } else {
      this.close();
    }
  },

  async open() {
    this._panel.classList.remove('hidden');
    if (!this._loaded) {
      await this.fetch();
    } else {
      this._render();
    }
    this._filter.focus();
  },

  close() {
    this._panel.classList.add('hidden');
  },

  async fetch() {
    this._body.innerHTML = '<div class="dossier-loading">Chargement des dossiers...</div>';
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=lister`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && Array.isArray(result.data)) {
        this._dossiers = result.data;
        this._loaded = true;
        this._render();
      } else {
        this._body.innerHTML = '<div class="dossier-empty">Erreur de chargement.</div>';
      }
    } catch (err) {
      this._body.innerHTML = '<div class="dossier-empty">Erreur de connexion.</div>';
      Toast.error('Impossible de charger la liste : ' + err.message);
    }
  },

  async refresh() {
    this._loaded = false;
    await this.fetch();
  },

  _render() {
    const query = (this._filter?.value || '').toLowerCase().trim();
    const filtered = query
      ? this._dossiers.filter(d =>
          d.ref.toLowerCase().includes(query) ||
          d.client.toLowerCase().includes(query)
        )
      : this._dossiers;

    if (filtered.length === 0) {
      this._body.innerHTML = `<div class="dossier-empty">${query ? 'Aucun résultat pour "' + query + '"' : 'Aucun dossier trouvé.'}</div>`;
      return;
    }

    const fmtCHF = (v) => {
      const n = parseFloat(v) || 0;
      return n.toLocaleString('fr-CH', { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + ' CHF';
    };

    this._body.innerHTML = filtered.map(d => `
      <div class="dossier-row" data-ref="${d.ref}">
        <div>
          <div class="ref">${d.ref}</div>
          <div class="client">${d.client || '—'}</div>
        </div>
        <div class="montant">${fmtCHF(d.montantHT)}</div>
        <div class="meta">${d.date || ''}</div>
      </div>
    `).join('');

    this._body.querySelectorAll('.dossier-row').forEach(row => {
      row.addEventListener('click', () => {
        const ref = row.dataset.ref;
        document.getElementById('searchRef').value = ref;
        this.close();
        DossierLoader.load(ref);
      });
    });
  }
};

// ============================================
// DOSSIER LOADER (via fetch)
// ============================================
const DossierLoader = {
  async load(ref) {
    if (!ref.trim()) { Toast.warning('Entrez une référence.'); return; }

    const btn = document.getElementById('loadBtn');
    btn.disabled = true;
    btn.textContent = '...';

    try {
      const url = `${CONFIG.SCRIPT_URL}?action=rechercher&ref=${encodeURIComponent(ref.trim())}`;
      const resp = await fetch(url);
      const result = await resp.json();

      if (!result || result.status === 'not_found') {
        Toast.warning('Dossier introuvable.');
        return;
      }

      const data = result.data || result;
      const form = document.getElementById('devisForm');
      if (data.ref) form.querySelector('[name="ref"]').value = data.ref;
      if (data.client) form.querySelector('[name="client"]').value = data.client;
      if (data.adresseClient) form.querySelector('[name="adresseClient"]').value = data.adresseClient;
      if (data.adresseDepart) form.querySelector('[name="adresseDepart"]').value = data.adresseDepart;
      if (data.adresseArrivee) form.querySelector('[name="adresseArrivee"]').value = data.adresseArrivee;
      if (data.contact) form.querySelector('[name="contact"]').value = data.contact;
      if (data.contactTel) form.querySelector('[name="contactTel"]').value = data.contactTel;
      if (data.montantHT) form.querySelector('[name="montantHT"]').value = data.montantHT;
      if (data.genre) form.querySelector('[name="genre"]').value = data.genre;
      if (data.volumeEstime) form.querySelector('[name="volumeEstime"]').value = data.volumeEstime;

      PosteManager.loadPostes(data.postes);
      PriceCalc.updateBreakdown();
      Toast.success('Dossier chargé !');
    } catch (err) {
      Toast.error('Erreur de chargement : ' + err.message);
    } finally {
      btn.disabled = false;
      btn.textContent = 'Ouvrir';
    }
  }
};

// ============================================
// FORM SUBMITTER
// ============================================
const FormSubmitter = {
  _submitting: false,

  async submit(form) {
    if (this._submitting) return;
    if (!Validator.validate(form)) {
      Toast.error('Corrigez les erreurs avant de soumettre.');
      return;
    }

    const fd = new FormData(form);
    const data = Object.fromEntries(fd.entries());
    data.postes = PosteManager.collectAll();

    Modal.show(data, () => this._doSubmit(data));
  },

  async _doSubmit(data) {
    this._submitting = true;
    const btn = document.getElementById('submitBtn');
    const btnText = document.getElementById('btnText');
    const loader = document.getElementById('loader');

    btn.disabled = true;
    btnText.textContent = 'Envoi en cours...';
    loader.classList.remove('hidden');

    try {
      const response = await fetch(CONFIG.SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify(data)
      });
      const res = await response.json();

      if (res.status === 'success') {
        Toast.success('Devis généré avec succès !');
        if (res.resa === 'ok') {
          Toast.success('Fiche RESA créée !');
        } else if (res.resa === 'error') {
          Toast.error('Erreur RESA : ' + (res.resaError || 'Inconnue'));
        }
        const waNumber = (data.mobileSociete || '').replace(/\s+/g, '');
        const waText = encodeURIComponent(`Documents ${data.ref}:\nPDF: ${res.pdfUrl}`);
        window.open(`https://wa.me/${waNumber}?text=${waText}`, '_blank');
      } else {
        Toast.error('Erreur serveur : ' + (res.message || 'Inconnue'));
      }
    } catch (err) {
      Toast.error('Erreur de connexion : ' + err.message);
    } finally {
      this._submitting = false;
      btn.disabled = false;
      btnText.textContent = 'Générer Devis + RESA';
      loader.classList.add('hidden');
    }
  }
};

// ============================================
// SYNC LOGIC
// ============================================
const SyncManager = {
  init() {
    const vTel = document.getElementById('vendeurTel');
    const waMob = document.getElementById('mobileSociete');
    let autoSyncWA = true;
    if (vTel && waMob) {
      vTel.addEventListener('input', () => { if (autoSyncWA) waMob.value = vTel.value; });
      waMob.addEventListener('input', () => { autoSyncWA = false; });
    }

    const factAdr = document.getElementById('adresseFacturation');
    const depAdr = document.getElementById('adresseDepart');
    let autoSyncDep = true;
    if (factAdr && depAdr) {
      factAdr.addEventListener('input', () => { if (autoSyncDep) depAdr.value = factAdr.value; });
      depAdr.addEventListener('input', () => { autoSyncDep = false; });
    }
  }
};

// ============================================
// PELICHET / AUTRE TOGGLE
// ============================================
function togglePelichet(isPelichet) {
  document.querySelectorAll('.pelichet-only').forEach(el => {
    el.style.display = isPelichet ? '' : 'none';
  });
}

// ============================================
// INIT
// ============================================
document.addEventListener('DOMContentLoaded', async () => {
  // Charger les données depuis le serveur
  await Promise.all([
    CustomItems.fetchFromServer(),
    TarifManager.fetchFromServer()
  ]);

  SyncManager.init();
  PosteManager.init();
  PriceCalc.init();
  DossierList.init();
  SettingsPanel.init();

  // Search bar
  document.getElementById('loadBtn')?.addEventListener('click', () => {
    DossierLoader.load(document.getElementById('searchRef').value);
  });
  document.getElementById('searchRef')?.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); DossierLoader.load(e.target.value); }
  });

  // Auto total button
  document.getElementById('autoTotalBtn')?.addEventListener('click', () => PriceCalc.applyAutoTotal());

  // Form submit
  document.getElementById('devisForm')?.addEventListener('submit', (e) => {
    e.preventDefault();
    FormSubmitter.submit(e.target);
  });
});
