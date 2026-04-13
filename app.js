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

  _createOptionsHTML(list) {
    return '<option value="">-- Type --</option>'
      + list.map(i => `<option value="${i}">${i}</option>`).join('')
      + '<option value="__autre__">Autre (personnalisé)</option>';
  },

  _createRow(type) {
    const isEngin = type === 'engins';
    const div = document.createElement('div');
    div.className = 'row-item';
    div.dataset.type = type;
    div.innerHTML = `
      <select style="flex:1">${this._createOptionsHTML(CONFIG.LISTS[type])}</select>
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
    // Toggle champ personnalisé quand "Autre" est sélectionné
    const select = div.querySelector('select');
    const customInput = div.querySelector('.custom-input');
    select.addEventListener('change', () => {
      if (select.value === '__autre__') {
        customInput.classList.remove('hidden');
        customInput.focus();
      } else {
        customInput.classList.add('hidden');
        customInput.value = '';
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
        </div>
      `}
    `;

    // Remove button
    div.querySelector('.poste-remove').addEventListener('click', () => {
      div.remove();
      this.renumber();
      PriceCalc.updateBreakdown();
    });

    // Prix change -> update breakdown
    const prixInput = div.querySelector('[name="postePrix"]');
    if (prixInput) prixInput.addEventListener('input', () => PriceCalc.updateBreakdown());

    // Detail mode event bindings
    if (mode === 'detail') {
      const rdvContainer = div.querySelector('.rdv-container');
      rdvContainer.appendChild(this._createRDVRow());

      div.querySelector('.add-row-btn.rdv').addEventListener('click', () => {
        rdvContainer.appendChild(this._createRDVRow());
      });

      const nbJoursInput = div.querySelector('[name="nbJours"]');
      div.querySelector('[data-half]').addEventListener('click', () => { nbJoursInput.value = '0.5'; });
      div.querySelector('[data-one]').addEventListener('click', () => { nbJoursInput.value = '1'; });
      div.querySelector('[data-two]').addEventListener('click', () => { nbJoursInput.value = '2'; });

      ['engins', 'personnel', 'vehicules', 'materiel'].forEach(type => {
        const section = div.querySelector(`.detail-section.${type}`);
        section.querySelector('.add-row-btn').addEventListener('click', () => {
          section.querySelector('.rows').appendChild(this._createRow(type));
        });
      });
    }

    this._container.appendChild(div);
    this.renumber();
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
      this.addPoste(p.mode || 'simple');
      const card = this._container.lastElementChild;
      card.querySelector('[name="posteTitre"]').value = p.titre || '';
      card.querySelector('[name="postePrix"]').value = p.prix || '';

      if (p.mode === 'simple') {
        const ta = card.querySelector('[name="simpleText"]');
        if (ta) ta.value = p.text || '';
      }
    });
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
document.addEventListener('DOMContentLoaded', () => {
  SyncManager.init();
  PosteManager.init();
  PriceCalc.init();
  DossierList.init();

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
