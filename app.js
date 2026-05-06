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
    materiel: ['Chariots', 'Couvertures de laine', 'Rouleaux bulle', 'Adhésif', 'Transpalette', 'Transpalette electrique', 'Gerbeur electrique', 'Caisses à outils', 'Sangles', 'Diable']
  }
};

// ============================================
// USER MANAGER (login multi-utilisateur)
// ============================================
const UserManager = {
  _currentUser: null,
  _users: [],

  /** Charge les utilisateurs depuis le serveur */
  async fetchUsers() {
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=users_get`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && Array.isArray(result.data)) {
        this._users = result.data;
        localStorage.setItem('pelichet_users', JSON.stringify(this._users));
      }
    } catch (e) {
      console.log('UserManager fetch failed, using localStorage');
      try {
        const raw = localStorage.getItem('pelichet_users');
        if (raw) this._users = JSON.parse(raw);
      } catch (err) { /* silencieux */ }
    }
  },

  /** Affiche l'ecran de login */
  showLogin() {
    return new Promise((resolve) => {
      const overlay = document.getElementById('loginOverlay');
      const grid = document.getElementById('loginUsersGrid');
      const loginView = document.getElementById('loginView');
      const registerView = document.getElementById('registerView');
      if (!overlay || !grid) { resolve(null); return; }

      // Stocker le resolve pour l'utiliser depuis register
      this._loginResolve = resolve;

      // Verifier si un user est deja sauvegarde
      const savedId = localStorage.getItem('pelichet_current_user');
      if (savedId) {
        const saved = this._users.find(u => u.id === savedId);
        if (saved) {
          this._currentUser = saved;
          this._applyUser();
          overlay.classList.add('hidden');
          resolve(saved);
          return;
        }
      }

      // Toujours afficher la vue login par defaut
      if (loginView) loginView.classList.remove('hidden');
      if (registerView) registerView.classList.add('hidden');

      this._renderUserCards(resolve);
      overlay.classList.remove('hidden');

      // Bouton "Creer un compte" -> basculer vers le formulaire
      document.getElementById('showRegisterBtn')?.addEventListener('click', () => {
        if (loginView) loginView.classList.add('hidden');
        if (registerView) registerView.classList.remove('hidden');
        this._initSignatureUpload();
      });

      // Bouton "Retour" -> revenir au login
      document.getElementById('backToLoginBtn')?.addEventListener('click', () => {
        if (registerView) registerView.classList.add('hidden');
        if (loginView) loginView.classList.remove('hidden');
        document.getElementById('registerError')?.classList.add('hidden');
      });

      // Formulaire d'inscription
      document.getElementById('registerForm')?.addEventListener('submit', (e) => {
        e.preventDefault();
        this._handleRegister(resolve);
      });
    });
  },

  /** Genere les cartes utilisateurs dans la grille */
  _renderUserCards(resolve) {
    const grid = document.getElementById('loginUsersGrid');
    if (!grid) return;

    grid.innerHTML = this._users.map(u => `
      <button type="button" class="login-user-card" data-uid="${u.id}">
        <div class="login-avatar">${(u.prenom || u.nom || '?')[0].toUpperCase()}</div>
        <div class="login-info">
          <div class="login-name">${u.prenom} ${u.nom}</div>
          <div class="login-role">${u.titre || u.role || 'vendeur'}</div>
        </div>
      </button>
    `).join('');

    if (this._users.length === 0) {
      grid.innerHTML = '<div style="text-align:center;color:var(--slate-400);padding:1.5rem;font-size:0.85rem">Aucun compte.<br>Creez votre premier compte ci-dessous.</div>';
    }

    // Ecouter les clics
    grid.querySelectorAll('.login-user-card').forEach(card => {
      card.addEventListener('click', () => {
        const uid = card.dataset.uid;
        const user = this._users.find(u => u.id === uid);
        if (user) {
          this._currentUser = user;
          localStorage.setItem('pelichet_current_user', uid);
          this._applyUser();
          document.getElementById('loginOverlay')?.classList.add('hidden');
          Toast.success('Connecte : ' + user.prenom + ' ' + user.nom);
          resolve(user);
        }
      });
    });
  },

  /** Convertit un fichier image en base64 */
  _fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  },

  /** Initialise la zone de preview de signature */
  _initSignatureUpload() {
    const fileInput = document.getElementById('regSignature');
    const preview = document.getElementById('signaturePreview');
    if (!fileInput || !preview) return;

    fileInput.addEventListener('change', () => {
      const file = fileInput.files[0];
      if (!file) return;
      const url = URL.createObjectURL(file);
      preview.innerHTML = `<img src="${url}" alt="Signature">`;
      preview.classList.add('has-image');
    });

    // Drag & drop
    const zone = document.getElementById('signatureZone');
    if (zone) {
      zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
      zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file && file.type.startsWith('image/')) {
          // Mettre le fichier dans l'input
          const dt = new DataTransfer();
          dt.items.add(file);
          fileInput.files = dt.files;
          const url = URL.createObjectURL(file);
          preview.innerHTML = `<img src="${url}" alt="Signature">`;
          preview.classList.add('has-image');
        }
      });
    }
  },

  /** Gere la soumission du formulaire d'inscription */
  async _handleRegister(resolve) {
    const prenom = document.getElementById('regPrenom')?.value?.trim();
    const nom = document.getElementById('regNom')?.value?.trim();
    const telephone = document.getElementById('regTelephone')?.value?.trim();
    const email = document.getElementById('regEmail')?.value?.trim();
    const role = document.getElementById('regRole')?.value || 'vendeur';
    const titre = document.getElementById('regTitre')?.value?.trim();
    const signatureFile = document.getElementById('regSignature')?.files?.[0];
    const errorDiv = document.getElementById('registerError');
    const submitBtn = document.getElementById('registerSubmitBtn');

    // Validation
    if (!prenom || !nom) {
      if (errorDiv) { errorDiv.textContent = 'Prenom et nom sont obligatoires.'; errorDiv.classList.remove('hidden'); }
      return;
    }
    if (!telephone) {
      if (errorDiv) { errorDiv.textContent = 'Le numero de telephone est obligatoire.'; errorDiv.classList.remove('hidden'); }
      return;
    }
    if (!titre) {
      if (errorDiv) { errorDiv.textContent = 'Le titre / fonction est obligatoire.'; errorDiv.classList.remove('hidden'); }
      return;
    }

    // Desactiver le bouton
    if (submitBtn) { submitBtn.disabled = true; submitBtn.textContent = 'Creation en cours...'; }
    if (errorDiv) errorDiv.classList.add('hidden');

    try {
      // Convertir la signature en base64 si presente
      let signatureBase64 = '';
      if (signatureFile) {
        signatureBase64 = await this._fileToBase64(signatureFile);
      }

      const userData = { _action: 'user_add', prenom, nom, telephone, email, role, titre, signatureBase64 };
      const resp = await fetch(CONFIG.SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify(userData)
      });
      const result = await resp.json();

      if (result.status === 'success' && result.data) {
        const newUser = result.data;
        // Ajouter a la liste locale
        this._users.push(newUser);
        localStorage.setItem('pelichet_users', JSON.stringify(this._users));

        // Connecter directement le nouvel utilisateur
        this._currentUser = newUser;
        localStorage.setItem('pelichet_current_user', newUser.id);
        this._applyUser();

        document.getElementById('loginOverlay')?.classList.add('hidden');
        Toast.success('Compte cree ! Bienvenue ' + newUser.prenom);

        // Reset formulaire
        document.getElementById('registerForm')?.reset();

        resolve(newUser);
      } else {
        if (errorDiv) { errorDiv.textContent = result.message || 'Erreur lors de la creation.'; errorDiv.classList.remove('hidden'); }
      }
    } catch (err) {
      if (errorDiv) { errorDiv.textContent = 'Erreur de connexion : ' + err.message; errorDiv.classList.remove('hidden'); }
    } finally {
      if (submitBtn) { submitBtn.disabled = false; submitBtn.textContent = 'Creer mon compte'; }
    }
  },

  /** Applique les infos utilisateur au formulaire */
  _applyUser() {
    if (!this._currentUser) return;
    const u = this._currentUser;
    const fullName = (u.prenom + ' ' + u.nom).toUpperCase();

    // Nettoyer le telephone (ignorer #ERROR!, #REF!, etc.)
    const tel = (u.telephone && !u.telephone.startsWith('#')) ? u.telephone : '';

    // Mettre a jour le nav
    const navUser = document.getElementById('navUserName');
    if (navUser) navUser.textContent = u.prenom;

    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) logoutBtn.classList.remove('hidden');

    const profileBtn = document.getElementById('profileBtn');
    if (profileBtn) profileBtn.classList.remove('hidden');

    // Pre-remplir vendeur
    const vendeur = document.getElementById('vendeur');
    const vendeurTel = document.getElementById('vendeurTel');
    if (vendeur) vendeur.value = fullName;
    if (vendeurTel) vendeurTel.value = tel;

    // Pre-remplir WhatsApp
    const waMob = document.getElementById('mobileSociete');
    if (waMob && tel) waMob.value = tel;
  },

  /** Deconnexion */
  logout() {
    this._currentUser = null;
    localStorage.removeItem('pelichet_current_user');
    const navUser = document.getElementById('navUserName');
    if (navUser) navUser.textContent = '';
    const logoutBtn = document.getElementById('logoutBtn');
    if (logoutBtn) logoutBtn.classList.add('hidden');

    // Reafficher le login
    const overlay = document.getElementById('loginOverlay');
    const loginView = document.getElementById('loginView');
    const registerView = document.getElementById('registerView');
    if (loginView) loginView.classList.remove('hidden');
    if (registerView) registerView.classList.add('hidden');
    if (overlay) overlay.classList.remove('hidden');

    // Re-render les cartes (avec un nouveau resolve)
    this._renderUserCards((user) => {
      // Callback quand un user est selectionne apres logout
    });
  },

  /** Retourne l'ID de l'utilisateur connecte */
  getUserId() {
    return this._currentUser ? this._currentUser.id : '';
  },

  /** Retourne l'utilisateur connecte */
  getUser() {
    return this._currentUser;
  }
};

// ============================================
// KM CALCULATOR (calcul kilométrique)
// ============================================
const KmCalculator = {
  // Fallback CHF/km par véhicule si non défini dans les tarifs (grille Pelichet historique)
  // Les valeurs réelles sont lues depuis Settings → Tarifs (colonne CHF/km)
  KM_RATES_FALLBACK: {
    'VL': 0.8, '1F': 0.8, 'Box IT': 0.8, 'PL avec hayon': 1.5,
    'PL': 1.5, 'Semi': 1.7
  },
  // Catégories pour l'affichage groupé (informatif)
  KM_CATEGORIES: [
    { label: 'Fourgon / Voiture (VL, 1F, Box IT)', rate: 0.8 },
    { label: 'Camion 18T (PL, PL avec hayon)', rate: 1.5 },
    { label: 'Tracteur Remorque (Semi)', rate: 1.7 }
  ],

  /** Retourne le tarif CHF/km d'un véhicule depuis Settings (avec fallback) */
  _getRate(vehName) {
    // Priorité 1 : tarif défini dans Settings (vehicules)
    let r = TarifManager?.getCoutKm?.('vehicules', vehName) || 0;
    if (r > 0) return r;
    // Priorité 2 : tarif défini sur les engins (parfois utilisés comme véhicules spéciaux)
    r = TarifManager?.getCoutKm?.('engins', vehName) || 0;
    if (r > 0) return r;
    // Priorité 3 : fallback historique
    return this.KM_RATES_FALLBACK[vehName] || 0;
  },

  // Compatibilité descendante : KM_RATES utilisé partout via getter
  get KM_RATES() {
    return new Proxy({}, {
      get: (_, key) => this._getRate(key)
    });
  },
  FORFAIT_KM: 50,
  _lastResult: null,

  init() {
    document.getElementById('kmCalcBtn')?.addEventListener('click', () => this.calculate());
  },

  async calculate() {
    const depart = document.querySelector('[name="adresseDepart"]')?.value?.trim();
    const arrivee = document.querySelector('[name="adresseArrivee"]')?.value?.trim();

    if (!depart || !arrivee) {
      Toast.warning('Remplissez les adresses de départ et d\'arrivée.');
      return;
    }

    const btn = document.getElementById('kmCalcBtn');
    btn.classList.add('loading');
    btn.textContent = '⏳ Calcul en cours...';

    try {
      const url = `${CONFIG.SCRIPT_URL}?action=km_calc&depart=${encodeURIComponent(depart)}&arrivee=${encodeURIComponent(arrivee)}`;
      const resp = await fetch(url);
      const result = await resp.json();

      if (result.status !== 'success' || !result.data) {
        Toast.error('Erreur : ' + (result.message || 'Calcul impossible'));
        return;
      }

      this._lastResult = result.data;
      this._showResult(result.data);
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    } finally {
      btn.classList.remove('loading');
      btn.textContent = '🛣️ Calculer Kilométrage';
    }
  },

  _showResult(data) {
    const div = document.getElementById('kmResult');
    const excessClass = data.excessKm > 0 ? 'has-excess' : 'zero';
    const excessText = data.excessKm > 0
      ? `${data.excessKm} km supplémentaires (au-delà du forfait de ${data.forfaitKm} km)`
      : `Aucun supplément (${data.totalKm} km ≤ ${data.forfaitKm} km forfait)`;

    // Calculer les coûts par véhicule défini dans Settings
    let costsHTML = '';
    let addBtnHTML = '';
    if (data.excessKm > 0) {
      const tarifsVeh = (TarifManager.getData()?.items?.vehicules || [])
        .filter(v => parseFloat(v.coutKm) > 0);
      costsHTML = '<div class="km-costs"><div style="font-weight:600;margin-bottom:4px">Coût supplément / véhicule (par km) :</div>';
      if (tarifsVeh.length === 0) {
        // Fallback : afficher les catégories si aucun véhicule n'a coutKm défini
        this.KM_CATEGORIES.forEach(cat => {
          const cost = (data.excessKm * cat.rate).toFixed(2);
          costsHTML += `<div class="km-cost-row"><span>${cat.label}</span><span class="cost-val">${cost} CHF</span></div>`;
        });
        costsHTML += '<div class="km-cost-row" style="font-style:italic;color:var(--ink-4);margin-top:4px;font-size:10px">Astuce : définis CHF/km dans ⚙️ Tarifs pour personnaliser</div>';
      } else {
        tarifsVeh.forEach(v => {
          const cost = (data.excessKm * parseFloat(v.coutKm)).toFixed(2);
          costsHTML += `<div class="km-cost-row"><span>${v.item} (${v.coutKm} CHF/km)</span><span class="cost-val">${cost} CHF</span></div>`;
        });
      }
      costsHTML += '</div>';

      // Calculer le coût total en fonction des véhicules déjà sélectionnés dans les postes
      const totalKmCost = this._calcTotalKmCost(data.excessKm);
      addBtnHTML = `
        <button type="button" class="btn km-add-poste-btn" id="kmAddPosteBtn" style="margin-top:0.75rem;width:100%;background:var(--blue);color:white;font-size:0.75rem;padding:0.5rem">
          ➕ Ajouter frais km au devis (${totalKmCost.toFixed(2)} CHF)
        </button>
      `;
    }

    // Détails des segments
    let legsHTML = '<div class="km-legs">';
    (data.legs || []).forEach(leg => {
      legsHTML += `<div class="km-leg"><span>${leg.from.split(',')[0]} → ${leg.to.split(',')[0]}</span><span>${leg.km} km (${leg.duration})</span></div>`;
    });
    legsHTML += '</div>';

    div.innerHTML = `
      <div class="km-total">📍 ${data.totalKm} km total (aller-retour)</div>
      <div class="km-excess ${excessClass}">${excessText}</div>
      ${legsHTML}
      ${costsHTML}
      ${addBtnHTML}
    `;
    div.classList.remove('hidden');

    // Bind click sur le bouton "Ajouter au devis"
    document.getElementById('kmAddPosteBtn')?.addEventListener('click', () => this.ajouterPosteKm());
  },

  /** Convertit une valeur de durée en nombre de jours */
  _dureeEnJours(dureeVal, nbJoursMission) {
    if (!dureeVal) return nbJoursMission || 1;
    if (dureeVal === '1/2 AM' || dureeVal === '1/2 PM') return 0.5;
    const m = dureeVal.match(/^(\d+)j$/);
    return m ? parseInt(m[1]) : (nbJoursMission || 1);
  },

  /** Collecte tous les véhicules avec qty × jours (aggrégé par type de véhicule) */
  _collectVehiculesJours() {
    const result = {}; // { 'PL': { qtyJours: 4, details: [...] } }
    const postes = PosteManager.collectAll();

    postes.forEach((p, posteIdx) => {
      const nbJoursPoste = parseFloat(p.jours) || 1;
      const titrePoste = (p.titre || `Poste ${posteIdx + 1}`).toUpperCase();

      (p.vehicules || []).forEach(v => {
        const m = String(v).match(/^(\d+)x\s+(.+?)(?:\s+\[(.+?)\])?$/);
        if (!m) return;
        const qty = parseInt(m[1]) || 1;
        const veh = m[2].trim();
        const duree = m[3] || '';
        const jours = this._dureeEnJours(duree, nbJoursPoste);
        const qtyJours = qty * jours;

        if (!result[veh]) result[veh] = { qtyJours: 0, lines: [] };
        result[veh].qtyJours += qtyJours;
        result[veh].lines.push({
          qty, jours, duree: duree || `${nbJoursPoste}j (mission)`, poste: titrePoste
        });
      });
    });

    return result;
  },

  /** Calcule le coût total km : excessKm × rate × (qty × jours) pour chaque véhicule */
  _calcTotalKmCost(excessKm) {
    let total = 0;
    const vehMap = this._collectVehiculesJours();
    Object.entries(vehMap).forEach(([veh, info]) => {
      const rate = this.KM_RATES[veh];
      if (rate) total += info.qtyJours * excessKm * rate;
    });
    return total;
  },

  /** Distribue les frais km sur les postes existants (ajoute aux prix, pas de poste séparé) */
  ajouterPosteKm() {
    if (!this._lastResult || this._lastResult.excessKm <= 0) {
      Toast.warning('Pas de frais kilométriques à ajouter.');
      return;
    }

    const excessKm = this._lastResult.excessKm;
    const cards = document.querySelectorAll('#prestationsContainer .poste-card');

    if (cards.length === 0) {
      Toast.warning('Ajoutez d\'abord un poste avant de calculer les frais km.');
      return;
    }

    let totalAjoute = 0;

    cards.forEach(card => {
      const nbJoursPoste = parseFloat(card.querySelector('[name="nbJours"]')?.value) || 1;

      // Calculer le coût km pour ce poste (somme des véh-jours × excessKm × rate)
      let kmCostPoste = 0;
      const vGrid = card.querySelector('.cb-grid[data-type="vehicules"]');
      if (!vGrid) return;

      vGrid.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
        const veh = cb.value;
        const rate = this.KM_RATES[veh];
        if (!rate) return;

        const cbItem = cb.closest('.cb-item');
        const rows = cbItem.querySelectorAll('.cb-row');
        if (rows.length > 0) {
          rows.forEach(row => {
            const qty = parseFloat(row.querySelector('.cb-qty')?.value) || 1;
            const duree = row.querySelector('.cb-duree')?.value || '';
            const jours = this._dureeEnJours(duree, nbJoursPoste);
            kmCostPoste += qty * jours * excessKm * rate;
          });
        }
      });

      if (kmCostPoste <= 0) return;

      // Récupérer le prix de base (stocké dans dataset, ou prix actuel si première application)
      const prixInput = card.querySelector('[name="postePrix"]');
      if (!prixInput) return;

      let prixBase = parseFloat(prixInput.dataset.prixBase);
      if (isNaN(prixBase)) {
        prixBase = parseFloat(prixInput.value) || 0;
        prixInput.dataset.prixBase = prixBase.toString();
      }

      // Nouveau prix = prix base + frais km
      const nouveauPrix = prixBase + kmCostPoste;
      prixInput.value = nouveauPrix.toFixed(2);
      prixInput.dataset.manual = 'true';
      prixInput.dataset.kmCost = kmCostPoste.toFixed(2);

      totalAjoute += kmCostPoste;
    });

    if (totalAjoute === 0) {
      Toast.warning('Aucun véhicule facturable km dans les postes (vérifie qu\'ils ont un tarif km défini).');
      return;
    }

    // Recalcule automatiquement le total HT global à partir des prix postes
    PriceCalc.applyAutoTotal();
    PriceCalc.updateBreakdown();
    if (typeof RightRail !== 'undefined') RightRail.update();
    Toast.success(`Frais km ajoutés : +${totalAjoute.toFixed(2)} CHF · Total HT actualisé`);
  },

  /** Retourne les km supplémentaires (pour la ventilation) */
  getExcessKm() {
    return this._lastResult ? this._lastResult.excessKm : 0;
  },

  getTotalKm() {
    return this._lastResult ? this._lastResult.totalKm : 0;
  }
};

// ============================================
// COMPANY SEARCH (autocomplétion Zefix)
// ============================================
const CompanySearch = {
  _timer: null,
  _dropdown: null,
  _input: null,

  init() {
    this._input = document.querySelector('input[name="client"]');
    if (!this._input) return;

    // Créer le wrapper pour positionner le dropdown
    const wrapper = document.createElement('div');
    wrapper.className = 'company-search-wrap';
    this._input.parentNode.insertBefore(wrapper, this._input);
    wrapper.appendChild(this._input);
    // Déplacer l'error-msg dans le wrapper aussi
    const errMsg = wrapper.nextElementSibling;
    if (errMsg && errMsg.classList.contains('error-msg')) wrapper.appendChild(errMsg);

    // Créer le dropdown
    this._dropdown = document.createElement('div');
    this._dropdown.className = 'company-dropdown hidden';
    wrapper.appendChild(this._dropdown);

    // Événements
    this._input.addEventListener('input', () => this._onInput());
    this._input.addEventListener('focus', () => { if (this._dropdown.children.length > 0) this._dropdown.classList.remove('hidden'); });
    document.addEventListener('click', (e) => {
      if (!wrapper.contains(e.target)) this._dropdown.classList.add('hidden');
    });
  },

  _onInput() {
    clearTimeout(this._timer);
    const q = this._input.value.trim();
    if (q.length < 3) { this._dropdown.classList.add('hidden'); return; }
    this._timer = setTimeout(() => this._search(q), 600);
  },

  async _search(query) {
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=company_search&q=${encodeURIComponent(query)}`;
      const resp = await fetch(url);
      const data = await resp.json();
      if (data.status === 'success' && data.results?.length > 0) {
        this._showResults(data.results);
      } else {
        this._dropdown.classList.add('hidden');
      }
    } catch (e) {
      console.log('Company search error:', e);
    }
  },

  _showResults(results) {
    this._dropdown.innerHTML = '';
    results.forEach(r => {
      const cityLine = [r.zip, r.city].filter(Boolean).join(' ');
      const addrParts = [r.street, cityLine].filter(Boolean);
      const displayName = r.name || r.street || cityLine;
      const displayAddr = r.name ? addrParts.join(', ') : (addrParts.length > 1 ? addrParts.join(', ') : '');

      const item = document.createElement('div');
      item.className = 'company-result';
      item.innerHTML = `
        <div class="company-result-name">${this._esc(displayName)}</div>
        ${displayAddr ? `<div class="company-result-addr">${this._esc(displayAddr)}</div>` : ''}
      `;
      item.addEventListener('click', () => this._select(r));
      this._dropdown.appendChild(item);
    });
    this._dropdown.classList.remove('hidden');
  },

  _select(company) {
    // Remplir le nom (nom d'entreprise si dispo, sinon laisser tel quel)
    if (company.name) this._input.value = company.name;

    // Remplir l'adresse de facturation
    const cityLine = [company.zip, company.city].filter(Boolean).join(' ');
    const addr = [company.street, cityLine].filter(Boolean).join('\n');
    const addrField = document.querySelector('[name="adresseClient"]');
    if (addrField && addr) addrField.value = addr;

    this._dropdown.classList.add('hidden');
    Toast.info('Adresse renseignée');
  },

  _esc(str) {
    const d = document.createElement('div');
    d.textContent = str;
    return d.innerHTML;
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

  /** Retourne le coût demi-journée d'un item (si défini, sinon cout/2) */
  getCoutDemiJour(type, itemName) {
    if (!this._data || !this._data.items || !this._data.items[type]) return 0;
    const found = this._data.items[type].find(t => t.item === itemName);
    if (!found) return 0;
    // Si un prix demi-journée est défini, l'utiliser ; sinon moitié du prix journée
    return (found.coutDemiJour !== undefined && found.coutDemiJour !== null && found.coutDemiJour !== '')
      ? (parseFloat(found.coutDemiJour) || 0)
      : (parseFloat(found.cout) || 0) / 2;
  },

  /** Retourne le coût CHF/km d'un véhicule/engin (0 si non défini) */
  getCoutKm(type, itemName) {
    if (!this._data || !this._data.items || !this._data.items[type]) return 0;
    const found = this._data.items[type].find(t => t.item === itemName);
    if (!found) return 0;
    return parseFloat(found.coutKm) || 0;
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

  /** Vérifie si une durée est une demi-journée */
  _isDemiJour(dureeVal) {
    return dureeVal === '1/2 AM' || dureeVal === '1/2 PM';
  },

  /** Convertit une valeur de durée en nombre de jours */
  _dureeEnJours(dureeVal, nbJoursMission) {
    if (!dureeVal) return nbJoursMission; // durée mission par défaut
    if (this._isDemiJour(dureeVal)) return 0.5;
    const match = dureeVal.match(/^(\d+)j$/);
    return match ? parseInt(match[1]) : nbJoursMission;
  },

  /** Calcule le coût d'un item en fonction de sa durée */
  _coutItem(type, label, qty, dureeVal, nbJoursMission) {
    const unite = this.getUnite(type, label);
    if (unite !== 'jour') return qty * this.getCout(type, label); // pièce/forfait

    if (this._isDemiJour(dureeVal)) {
      // Utiliser le prix demi-journée spécifique (pas coût_jour × 0.5)
      return qty * this.getCoutDemiJour(type, label);
    }
    const jours = this._dureeEnJours(dureeVal, nbJoursMission);
    return qty * this.getCout(type, label) * jours;
  },

  /** Vérifie si un tarif existe pour un item ; renvoie true si OK, false sinon */
  _hasTariff(type, label) {
    if (!this._data?.items?.[type]) return false;
    return this._data.items[type].some(t => t.item === label);
  },

  /** Calcule le prix d'un poste détail à partir de ses ressources */
  calculerPrixPoste(card) {
    const nbJours = parseFloat(card.querySelector('[name="nbJours"]')?.value) || 1;
    let totalCout = 0;
    let totalPrix = 0;
    const details = [];
    const missing = []; // items sans tarif défini

    // --- 1. Engins (dropdown rows) ---
    const enginsRows = card.querySelector('.detail-section.engins .rows');
    let coutEngins = 0;
    if (enginsRows) {
      enginsRows.querySelectorAll('.row-item').forEach(row => {
        const selectVal = row.querySelector('select')?.value || '';
        const customVal = row.querySelector('input[name="customValue"]')?.value || '';
        const label = (selectVal === '__autre__') ? customVal : selectVal;
        if (!label) return;
        const qty = parseFloat(row.querySelector('input[name="qty"]')?.value) || 1;
        const dureeVal = row.querySelector('[name="enginDuree"]')?.value || '';
        const c = this._coutItem('engins', label, qty, dureeVal, nbJours);
        if (c === 0 && !this._hasTariff('engins', label)) missing.push({ type: 'engins', label });
        coutEngins += c;
      });
    }
    if (coutEngins > 0) {
      const marge = this.getMarge('engins');
      const prix = coutEngins * (1 + marge / 100);
      totalCout += coutEngins; totalPrix += prix;
      details.push({ type: 'engins', cout: coutEngins, marge, prix });
    }

    // --- 2. Personnel, Véhicules, Matériel (checkbox grids avec multi-lignes durée) ---
    ['personnel', 'vehicules', 'materiel'].forEach(type => {
      let coutCat = 0;
      const grid = card.querySelector(`.cb-grid[data-type="${type}"]`);
      if (!grid) return;

      grid.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
        const cbItem = cb.closest('.cb-item');
        const label = cb.value;
        const hasTariff = this._hasTariff(type, label);
        if (!hasTariff) missing.push({ type, label });

        const cbRows = cbItem.querySelectorAll('.cb-row');
        if (cbRows.length > 0) {
          cbRows.forEach(row => {
            const qty = parseFloat(row.querySelector('.cb-qty')?.value) || 1;
            const dureeVal = row.querySelector('.cb-duree')?.value || '';
            coutCat += this._coutItem(type, label, qty, dureeVal, nbJours);
          });
        } else {
          const qty = parseFloat(cbItem.querySelector('.cb-qty')?.value) || 1;
          coutCat += this._coutItem(type, label, qty, '', nbJours);
        }
      });

      if (coutCat > 0) {
        const marge = this.getMarge(type);
        const prix = coutCat * (1 + marge / 100);
        totalCout += coutCat; totalPrix += prix;
        details.push({ type, cout: coutCat, marge, prix });
      }
    });

    return {
      cout: Math.round(totalCout * 100) / 100,
      prix: Math.round(totalPrix * 100) / 100,
      details,
      missing
    };
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
        const showKm = (type === 'vehicules' || type === 'engins');
        tarifsHTML += `
          <tr data-type="${type}" data-idx="${idx}">
            <td><select class="tarif-type">${allTypes.map(tt => `<option value="${tt}" ${tt === type ? 'selected' : ''}>${tt}</option>`).join('')}</select></td>
            <td><input type="text" class="tarif-item" value="${t.item || ''}"></td>
            <td><input type="number" class="tarif-cout" value="${t.cout || 0}" min="0" step="0.5"></td>
            <td><input type="number" class="tarif-demi" value="${t.coutDemiJour ?? ''}" min="0" step="0.5" placeholder="auto"></td>
            <td><input type="number" class="tarif-km" value="${t.coutKm ?? ''}" min="0" step="0.1" placeholder="${showKm ? '0.0' : '—'}" ${showKm ? '' : 'disabled'}></td>
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
        <td><input type="number" class="tarif-demi" value="" min="0" step="0.5" placeholder="auto"></td>
        <td><input type="number" class="tarif-km" value="" min="0" step="0.1" placeholder="0.0"></td>
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
      const demiVal = tr.querySelector('.tarif-demi')?.value;
      const coutDemiJour = (demiVal !== undefined && demiVal !== '') ? parseFloat(demiVal) : null;
      const kmVal = tr.querySelector('.tarif-km')?.value;
      const coutKm = (kmVal !== undefined && kmVal !== '') ? parseFloat(kmVal) : null;
      const unite = tr.querySelector('.tarif-unite')?.value || 'jour';
      if (!type || !item) return;
      if (!data.items[type]) data.items[type] = [];
      const entry = { item, cout, unite };
      if (coutDemiJour !== null && !isNaN(coutDemiJour)) entry.coutDemiJour = coutDemiJour;
      if (coutKm !== null && !isNaN(coutKm)) entry.coutKm = coutKm;
      data.items[type].push(entry);
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
// PROFILE PANEL (edition profil utilisateur)
// ============================================
const ProfilePanel = {
  _panel: null,
  _open: false,

  init() {
    this._panel = document.getElementById('profilePanel');
    document.getElementById('profileBtn')?.addEventListener('click', () => this.toggle());
    document.getElementById('profileClose')?.addEventListener('click', () => this.close());
    document.getElementById('profileSave')?.addEventListener('click', () => this.save());
    this._initSignatureUpload();
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
    const u = UserManager.getUser();
    if (!u) return;

    document.getElementById('profPrenom').value = u.prenom || '';
    document.getElementById('profNom').value = u.nom || '';
    document.getElementById('profTelephone').value = u.telephone || '';
    document.getElementById('profTitre').value = u.titre || '';

    // Afficher la signature actuelle si elle existe
    const currentDiv = document.getElementById('profSignatureCurrent');
    const currentImg = document.getElementById('profSignatureImg');
    if (u.signatureId && currentDiv && currentImg) {
      currentImg.src = `https://drive.google.com/thumbnail?id=${u.signatureId}&sz=w400`;
      currentDiv.classList.remove('hidden');
    } else if (currentDiv) {
      currentDiv.classList.add('hidden');
    }

    // Reset le preview
    const preview = document.getElementById('profSignaturePreview');
    if (preview) {
      preview.innerHTML = '<span class="signature-placeholder">Cliquez ou glissez une nouvelle signature</span>';
      preview.classList.remove('has-image');
    }
    const fileInput = document.getElementById('profSignatureFile');
    if (fileInput) fileInput.value = '';
  },

  _initSignatureUpload() {
    const fileInput = document.getElementById('profSignatureFile');
    const preview = document.getElementById('profSignaturePreview');
    if (!fileInput || !preview) return;

    fileInput.addEventListener('change', () => {
      const file = fileInput.files[0];
      if (!file) return;
      const url = URL.createObjectURL(file);
      preview.innerHTML = `<img src="${url}" alt="Nouvelle signature">`;
      preview.classList.add('has-image');
    });

    // Drag & drop
    const zone = document.getElementById('profSignatureZone');
    if (zone) {
      zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
      zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
      zone.addEventListener('drop', (e) => {
        e.preventDefault();
        zone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file && file.type.startsWith('image/')) {
          const dt = new DataTransfer();
          dt.items.add(file);
          fileInput.files = dt.files;
          const url = URL.createObjectURL(file);
          preview.innerHTML = `<img src="${url}" alt="Nouvelle signature">`;
          preview.classList.add('has-image');
        }
      });
    }
  },

  async save() {
    const u = UserManager.getUser();
    if (!u) return;

    const titre = document.getElementById('profTitre')?.value?.trim() || '';
    const signatureFile = document.getElementById('profSignatureFile')?.files?.[0];
    const saveBtn = document.getElementById('profileSave');

    if (saveBtn) { saveBtn.disabled = true; saveBtn.textContent = 'Sauvegarde...'; }

    try {
      // 1. Mettre a jour le titre (via user_update_profile)
      const profileData = { userId: u.id, titre: titre };
      const urlProfile = `${CONFIG.SCRIPT_URL}?action=user_update_profile&data=${encodeURIComponent(JSON.stringify(profileData))}`;
      const respProfile = await fetch(urlProfile);
      const resultProfile = await respProfile.json();

      if (resultProfile.status === 'success') {
        // Mettre a jour le user local
        u.titre = titre;
        localStorage.setItem('pelichet_users', JSON.stringify(UserManager._users));
      }

      // 2. Upload signature si nouvelle image selectionnee
      if (signatureFile) {
        Toast.info('Upload de la signature...');
        const reader = new FileReader();
        const base64 = await new Promise((resolve, reject) => {
          reader.onload = () => resolve(reader.result);
          reader.onerror = reject;
          reader.readAsDataURL(signatureFile);
        });

        const respSig = await fetch(CONFIG.SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({ _action: 'user_upload_signature', userId: u.id, data: base64 })
        });
        const resultSig = await respSig.json();

        if (resultSig.status === 'success' && resultSig.signatureId) {
          u.signatureId = resultSig.signatureId;
          localStorage.setItem('pelichet_users', JSON.stringify(UserManager._users));
          Toast.success('Signature enregistree !');
        } else {
          Toast.error('Erreur signature : ' + (resultSig.message || 'Inconnue'));
        }
      }

      Toast.success('Profil mis a jour !');
      this.close();
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    } finally {
      if (saveBtn) { saveBtn.disabled = false; saveBtn.textContent = 'Sauvegarder'; }
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
    if (!this._htInput) return;
    const ht = parseFloat(this._htInput.value) || 0;
    const tva = ht * CONFIG.TVA_RATE;
    const rplp = ht * CONFIG.RPLP_RATE;
    const ttc = ht + tva + rplp;
    const fmt = (v) => v.toLocaleString('fr-CH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

    if (this._breakdownEl) {
      this._breakdownEl.innerHTML = `
        <div class="break-row"><span class="k">Total HT</span><span class="v">${fmt(ht)}</span></div>
        <div class="break-row"><span class="k">TVA (8.1%)</span><span class="v">${fmt(tva)}</span></div>
        <div class="break-row"><span class="k">RPLP (0.5%)</span><span class="v">${fmt(rplp)}</span></div>
        <div class="break-row tot"><span class="k">Total TTC</span><span class="v">${fmt(ttc)}</span></div>
      `;
      this._breakdownEl.classList.toggle('hidden', ht <= 0);
    }

    // Alimenter le right rail
    if (typeof RightRail !== 'undefined') RightRail.update();
  }
};

// ============================================
// CALENDAR VIEW (vue calendrier des missions)
// ============================================
const CalendarView = {
  _entries: [],
  _byDate: {},
  _cursor: null, // Premier jour du mois affiché
  _loaded: false,

  init() {
    document.getElementById('calendarBtn')?.addEventListener('click', () => this.open());
    document.getElementById('railCalendar')?.addEventListener('click', () => this.open());
    document.getElementById('calendarClose')?.addEventListener('click', () => this.close());
    document.getElementById('calendarPrev')?.addEventListener('click', () => this._nav(-1));
    document.getElementById('calendarNext')?.addEventListener('click', () => this._nav(+1));
    document.getElementById('calendarToday')?.addEventListener('click', () => { this._cursor = this._firstOfMonth(new Date()); this._render(); });
    document.getElementById('calendarSync')?.addEventListener('click', () => this._showSync());
    document.getElementById('calendarIcsDl')?.addEventListener('click', () => this._downloadIcs());
    document.getElementById('calendarSyncBannerBtn')?.addEventListener('click', () => this._showSync());
    document.getElementById('calendarSyncBannerClose')?.addEventListener('click', () => {
      document.getElementById('calendarSyncBanner')?.classList.add('hidden');
      localStorage.setItem('pelichet_cal_banner_hidden', '1');
    });
    if (localStorage.getItem('pelichet_cal_banner_hidden') === '1') {
      document.getElementById('calendarSyncBanner')?.classList.add('hidden');
    }

    document.getElementById('csmClose')?.addEventListener('click', () => document.getElementById('calendarSyncModal')?.classList.add('hidden'));
    document.getElementById('csmCopy')?.addEventListener('click', (e) => this._copyUrl(e.currentTarget));
    document.getElementById('csmDownloadIcs')?.addEventListener('click', () => this._downloadIcs());
    document.getElementById('csmOpenGoogle')?.addEventListener('click', () => {
      this._copyUrl();
      window.open('https://calendar.google.com/calendar/u/0/r/settings/addbyurl', '_blank');
    });
    document.getElementById('csmOpenOutlook')?.addEventListener('click', () => {
      this._copyUrl();
      window.open('https://outlook.live.com/calendar/0/addfromweb', '_blank');
    });

    // Onglets plateforme
    document.querySelectorAll('.csm-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.csm-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.csm-panel').forEach(p => p.classList.remove('active'));
        tab.classList.add('active');
        document.querySelector(`.csm-panel[data-panel="${tab.dataset.tab}"]`)?.classList.add('active');
      });
    });
    document.addEventListener('keydown', (e) => {
      if (!this._isOpen()) return;
      if (e.key === 'Escape') this.close();
      if (e.key === 'ArrowLeft') this._nav(-1);
      if (e.key === 'ArrowRight') this._nav(+1);
    });
    document.getElementById('calendarOverlay')?.addEventListener('click', (e) => {
      const detail = document.getElementById('calendarDetail');
      if (detail && !detail.classList.contains('hidden') && !detail.contains(e.target) && !e.target.closest('.cal-entry')) {
        detail.classList.add('hidden');
      }
    });
  },

  _isOpen() {
    return !document.getElementById('calendarOverlay')?.classList.contains('hidden');
  },

  async open() {
    document.getElementById('calendarOverlay').classList.remove('hidden');
    document.body.classList.add('modal-open');
    if (!this._cursor) this._cursor = this._firstOfMonth(new Date());
    if (!this._loaded) {
      await this._fetch();
    }
    this._render();
  },

  close() {
    document.getElementById('calendarOverlay').classList.add('hidden');
    document.getElementById('calendarDetail')?.classList.add('hidden');
    document.body.classList.remove('modal-open');
  },

  async _fetch() {
    const loading = document.getElementById('calendarLoading');
    loading?.classList.remove('hidden');
    try {
      const userId = UserManager.getUserId();
      const url = `${CONFIG.SCRIPT_URL}?action=calendar_data&user=${encodeURIComponent(userId)}`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && result.data) {
        this._entries = result.data.entries || [];
        this._byDate = {};
        this._entries.forEach(e => {
          if (!this._byDate[e.date]) this._byDate[e.date] = [];
          this._byDate[e.date].push(e);
        });
        this._loaded = true;
      } else {
        Toast.error('Erreur chargement calendrier : ' + (result.message || 'Inconnue'));
      }
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    } finally {
      loading?.classList.add('hidden');
    }
  },

  _nav(dir) {
    const d = new Date(this._cursor);
    d.setMonth(d.getMonth() + dir);
    this._cursor = this._firstOfMonth(d);
    this._render();
  },

  _firstOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); },

  _render() {
    if (!this._cursor) this._cursor = this._firstOfMonth(new Date());
    const year = this._cursor.getFullYear();
    const month = this._cursor.getMonth();

    // Titre
    const monthName = this._cursor.toLocaleDateString('fr-CH', { month: 'long', year: 'numeric' });
    document.getElementById('calendarTitle').textContent = monthName;

    // Déterminer le premier lundi à afficher (peut être dans le mois précédent)
    const first = new Date(year, month, 1);
    let startOffset = first.getDay() - 1; // Lundi = 1 → 0
    if (startOffset < 0) startOffset = 6; // Dimanche → fin
    const gridStart = new Date(year, month, 1 - startOffset);

    // 6 semaines × 7 jours = 42 cases
    const grid = document.getElementById('calendarGrid');
    const today = new Date();
    const todayStr = this._dateKey(today);

    const clientColors = {};
    let colorIdx = 0;
    const colorFor = (ref) => {
      if (clientColors[ref] === undefined) clientColors[ref] = (colorIdx++) % 6;
      return clientColors[ref];
    };

    let html = '';
    for (let i = 0; i < 42; i++) {
      const d = new Date(gridStart);
      d.setDate(gridStart.getDate() + i);
      const key = this._dateKey(d);
      const dayNum = d.getDate();
      const isOtherMonth = d.getMonth() !== month;
      const isWeekend = d.getDay() === 0 || d.getDay() === 6;
      const isToday = key === todayStr;

      const entries = this._byDate[key] || [];
      const shown = entries.slice(0, 3);
      const extra = entries.length - shown.length;

      const entriesHTML = shown.map(e => {
        const colorClass = `color-${colorFor(e.ref)}`;
        const personnelTags = (e.personnel || []).slice(0, 2).map(p => `<span class="cal-entry-res">${this._esc(p)}</span>`).join('');
        const vehTags = (e.vehicules || []).slice(0, 2).map(v => `<span class="cal-entry-res">${this._esc(v)}</span>`).join('');
        const enginTags = (e.engins || []).slice(0, 1).map(g => `<span class="cal-entry-res">${this._esc(g)}</span>`).join('');
        return `
          <div class="cal-entry ${colorClass}" data-date="${key}" data-ref="${this._esc(e.ref)}" data-idx="${entries.indexOf(e)}">
            <div class="cal-entry-client">${this._esc(e.client)}</div>
            <div class="cal-entry-resources">${personnelTags}${vehTags}${enginTags}</div>
          </div>
        `;
      }).join('');
      const moreHTML = extra > 0 ? `<div class="cal-day-more" data-date="${key}">+${extra} autre${extra > 1 ? 's' : ''}</div>` : '';

      const classes = ['cal-day'];
      if (isOtherMonth) classes.push('other-month');
      if (isWeekend) classes.push('weekend');
      if (isToday) classes.push('today');

      html += `
        <div class="${classes.join(' ')}" data-date="${key}">
          <span class="cal-day-num">${dayNum}</span>
          ${entriesHTML}
          ${moreHTML}
        </div>
      `;
    }

    grid.innerHTML = html;

    // Event : click entrée → detail popover
    grid.querySelectorAll('.cal-entry').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const date = el.dataset.date;
        const ref = el.dataset.ref;
        const entry = (this._byDate[date] || []).find(x => x.ref === ref);
        if (entry) this._showDetail(entry, el);
      });
    });

    // Event : click "+N autres" → detail du jour
    grid.querySelectorAll('.cal-day-more').forEach(el => {
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        const date = el.dataset.date;
        const entries = this._byDate[date] || [];
        this._showDayDetail(date, entries, el);
      });
    });
  },

  _showDetail(entry, anchor) {
    const detail = document.getElementById('calendarDetail');
    const dateFmt = new Date(entry.date).toLocaleDateString('fr-CH', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' });

    const list = (arr) => (arr || []).map(x => `<span class="cal-detail-tag">${this._esc(x)}</span>`).join('');
    const section = (title, items, badge) => {
      if (!items || items.length === 0) return '';
      return `
        <div class="cal-detail-section">
          <div class="cal-detail-section-title">${title} ${badge !== undefined ? `<span class="badge">${badge}</span>` : ''}</div>
          <div class="cal-detail-list">${list(items)}</div>
        </div>
      `;
    };

    detail.innerHTML = `
      <button type="button" class="cal-detail-close" aria-label="Fermer">×</button>
      <div class="cal-detail-date">${dateFmt}</div>
      <div class="cal-detail-client">${this._esc(entry.client)}</div>
      <div class="cal-detail-ref">${this._esc(entry.ref)}</div>
      <div class="cal-detail-titre">${this._esc(entry.titre)} — ${entry.jours}${entry.jours > 1 ? ' jours' : ' jour'}</div>
      ${section('Personnel', entry.personnel, entry.effectif || entry.personnel?.length)}
      ${section('Véhicules', entry.vehicules)}
      ${section('Véhicules spéciaux', entry.engins)}
      ${section('Matériel', entry.materiel)}
      ${entry.tache ? `
        <div class="cal-detail-section">
          <div class="cal-detail-section-title">À faire</div>
          <div class="cal-detail-tache">${this._esc(entry.tache)}</div>
        </div>
      ` : ''}
    `;

    detail.classList.remove('hidden');
    this._positionDetail(detail, anchor);

    detail.querySelector('.cal-detail-close').addEventListener('click', () => detail.classList.add('hidden'));
  },

  _showDayDetail(date, entries, anchor) {
    const detail = document.getElementById('calendarDetail');
    const dateFmt = new Date(date).toLocaleDateString('fr-CH', { weekday: 'long', day: 'numeric', month: 'long' });

    const cards = entries.map((e, i) => `
      <div style="padding:10px;background:var(--surface-2);border:1px solid var(--border);border-radius:var(--radius-sm);margin-bottom:8px;cursor:pointer"
           onclick="CalendarView._showDetail(CalendarView._byDate['${date}'][${i}], this)">
        <div style="font-weight:600;font-size:13px">${this._esc(e.client)}</div>
        <div style="font-size:11px;color:var(--ink-3);margin-top:2px">${this._esc(e.titre)} · ${e.effectif}H · ${(e.vehicules || []).join(', ') || '—'}</div>
      </div>
    `).join('');

    detail.innerHTML = `
      <button type="button" class="cal-detail-close" aria-label="Fermer">×</button>
      <div class="cal-detail-date">${dateFmt}</div>
      <div class="cal-detail-client" style="margin-bottom:12px">${entries.length} mission${entries.length > 1 ? 's' : ''}</div>
      ${cards}
    `;

    detail.classList.remove('hidden');
    this._positionDetail(detail, anchor);

    detail.querySelector('.cal-detail-close').addEventListener('click', () => detail.classList.add('hidden'));
  },

  _positionDetail(detail, anchor) {
    const rect = anchor.getBoundingClientRect();
    const detailW = 460;
    const vw = window.innerWidth;
    const vh = window.innerHeight;

    let left = rect.right + 8;
    if (left + detailW > vw - 20) left = rect.left - detailW - 8;
    if (left < 20) left = (vw - detailW) / 2; // fallback centré

    let top = rect.top;
    const detailH = detail.offsetHeight || 400;
    if (top + detailH > vh - 20) top = Math.max(20, vh - detailH - 20);

    detail.style.left = `${Math.max(20, left)}px`;
    detail.style.top = `${top}px`;

    // Mobile : centrer
    if (vw <= 768) {
      detail.style.left = '10px';
      detail.style.right = '10px';
      detail.style.top = '10px';
      detail.style.width = 'auto';
      detail.style.maxHeight = 'calc(100vh - 20px)';
    }
  },

  _dateKey(d) {
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${d.getFullYear()}-${mm}-${dd}`;
  },

  _icsUrl() {
    const userId = UserManager.getUserId();
    return `${CONFIG.SCRIPT_URL}?action=ics&user=${encodeURIComponent(userId)}`;
  },

  _showSync() {
    const modal = document.getElementById('calendarSyncModal');
    const input = document.getElementById('csmUrl');
    const url = this._icsUrl();
    if (input) input.value = url;
    // Lien webcal:// pour iPhone (ouvre directement l'app Calendrier)
    const webcal = url.replace(/^https?:\/\//, 'webcal://');
    const webcalLink = document.getElementById('csmWebcal');
    if (webcalLink) webcalLink.href = webcal;
    modal?.classList.remove('hidden');
  },

  _copyUrl(btn) {
    const input = document.getElementById('csmUrl');
    if (!input) return;
    const url = input.value;
    input.select();
    if (navigator.clipboard?.writeText) {
      navigator.clipboard.writeText(url).then(() => {
        Toast.success('URL copiée');
        if (btn) {
          const orig = btn.textContent;
          btn.textContent = '✓ Copié';
          setTimeout(() => { btn.textContent = orig; }, 1500);
        }
      }).catch(() => {
        document.execCommand('copy');
        Toast.success('URL copiée');
      });
    } else {
      document.execCommand('copy');
      Toast.success('URL copiée');
    }
  },

  async _downloadIcs() {
    try {
      const resp = await fetch(this._icsUrl());
      const text = await resp.text();
      const blob = new Blob([text], { type: 'text/calendar' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `pelichet-${UserManager.getUserId()}.ics`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      Toast.success('📅 Fichier .ics téléchargé');
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    }
  },

  _esc(s) {
    const d = document.createElement('div');
    d.textContent = String(s || '');
    return d.innerHTML;
  }
};

// ============================================
// RÉALISÉ MANAGER (saisie des ressources réelles jour par jour)
// ============================================
const RealiseManager = {
  _modal: null,
  _data: null, // structure { postes: [{titre, jours: [{date, personnel, vehicules, materiel, notes}]}], montantFacture, notes }

  init() {
    this._modal = document.getElementById('realiseModal');
    document.getElementById('realiseBtn')?.addEventListener('click', () => this.open());
    document.getElementById('realiseClose')?.addEventListener('click', () => this.close());
    document.getElementById('realiseSave')?.addEventListener('click', () => this.save());
    document.getElementById('realiseUseEstim')?.addEventListener('click', () => this._fillFromEstim());
    document.getElementById('realiseUseReel')?.addEventListener('click', () => this._fillFromReel());
  },

  open() {
    if (!this._modal) return;
    const ref = document.querySelector('[name="ref"]')?.value?.trim();
    if (!ref) { Toast.warning('Référence dossier requise.'); return; }

    const postes = PosteManager.collectAll().filter(p => p.mode === 'detail');
    if (postes.length === 0) { Toast.warning('Aucun poste détaillé à réaliser.'); return; }

    document.getElementById('realiseRefLabel').textContent = ref + ' · ' + (document.querySelector('[name="client"]')?.value || '');

    // Charger les données existantes : mémoire > localStorage > rien
    const existing = window._currentDossierRealise || this._restoreFromLocal(ref) || null;
    this._render(postes, existing);

    // Initialiser le montant facturé avec l'existant ou le HT du devis
    const inputFact = document.getElementById('realiseMontantFacture');
    const ht = parseFloat(document.querySelector('[name="montantHT"]')?.value) || 0;
    inputFact.value = (existing?.montantFacture ?? ht).toFixed(2);
    document.getElementById('realiseNotes').value = existing?.notes || '';

    this._modal.classList.remove('hidden');
    this._recompute();
  },

  close() { this._modal?.classList.add('hidden'); },

  _render(postes, existing) {
    const body = document.getElementById('realiseBody');
    const fmtDate = (d) => {
      if (!d) return '';
      const parts = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
      return parts ? `${parts[3]}.${parts[2]}.${parts[1]}` : d;
    };

    let html = '';
    postes.forEach((poste, pIdx) => {
      const titre = (poste.titre || `Poste ${pIdx + 1}`).toUpperCase();
      const nbJours = parseInt(poste.jours) || 1;
      // Construire la liste des jours à partir des RDVs et nbJours
      const jours = this._expandJours(poste, existing?.postes?.[pIdx]?.jours || []);

      // Estimés pour ce poste (utilisés en ref)
      const estPersonnel = (poste.personnel || []).join(', ') || '—';
      const estVehicules = [...(poste.vehicules || []), ...(poste.engins || [])].join(', ') || '—';
      const estMateriel = (poste.materiel || []).join(', ') || '—';

      html += `
        <div class="realise-poste" data-poste-idx="${pIdx}">
          <div class="realise-poste-head">
            <div>
              <div class="realise-poste-title">${this._esc(titre)}</div>
              <div class="realise-poste-sub">${jours.length} jour${jours.length > 1 ? 's' : ''} · ${this._esc(poste.tache || '')}</div>
            </div>
          </div>
      `;

      jours.forEach((j, jIdx) => {
        html += `
          <div class="realise-day" data-day-idx="${jIdx}">
            <div class="realise-day-head">
              <span class="realise-day-date">${fmtDate(j.date)}</span>
              <span class="realise-poste-sub">Jour ${jIdx + 1}/${jours.length}</span>
              <button type="button" class="realise-prefill-btn" data-poste-idx="${pIdx}" data-day-idx="${jIdx}" title="Pré-remplir avec l'estimé">⚡ Comme estimé</button>
            </div>
            <div class="realise-day-sections">
              <div class="realise-section">
                <div class="realise-section-head">
                  <span class="realise-section-title">Personnel réel</span>
                  <span class="realise-section-estim">Estim: ${this._esc(estPersonnel)}</span>
                </div>
                <div class="realise-chips" data-rtype="personnel" data-poste-idx="${pIdx}" data-day-idx="${jIdx}"></div>
              </div>
              <div class="realise-section">
                <div class="realise-section-head">
                  <span class="realise-section-title">Véhicules / Engins réel</span>
                  <span class="realise-section-estim">Estim: ${this._esc(estVehicules)}</span>
                </div>
                <div class="realise-chips" data-rtype="vehicules" data-poste-idx="${pIdx}" data-day-idx="${jIdx}"></div>
              </div>
              <div class="realise-section">
                <div class="realise-section-head">
                  <span class="realise-section-title">Matériel réel</span>
                  <span class="realise-section-estim">Estim: ${this._esc(estMateriel)}</span>
                </div>
                <div class="realise-chips" data-rtype="materiel" data-poste-idx="${pIdx}" data-day-idx="${jIdx}"></div>
              </div>
              <div class="realise-section" style="grid-column:1/-1">
                <label class="realise-field-label">Notes du jour</label>
                <input type="text" class="rj-notes" value="${this._esc(j.notes || '')}" placeholder="Remarques sur cette journée…">
              </div>
            </div>
          </div>
        `;
      });
      html += '</div>';
    });
    body.innerHTML = html;

    // Construire les chips pour chaque jour, pré-cochés depuis existing
    document.querySelectorAll('.realise-chips').forEach(grid => {
      const type = grid.dataset.rtype;
      const pIdx = parseInt(grid.dataset.posteIdx);
      const jIdx = parseInt(grid.dataset.dayIdx);
      const existingDay = (existing?.postes?.[pIdx]?.jours?.[jIdx]) || {};
      // Récupérer le contenu pré-existant (string format "2x X, 1x Y") OU le format checkbox tableau
      let preEntries = [];
      const raw = existingDay[type === 'vehicules' ? 'vehicules' : type] || '';
      if (Array.isArray(raw)) {
        preEntries = raw;
      } else if (typeof raw === 'string' && raw.trim()) {
        preEntries = this._parseEntries(raw);
      }
      this._buildRealiseChips(grid, type, preEntries);
    });

    // Boutons "Comme estimé"
    document.querySelectorAll('.realise-prefill-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const pIdx = parseInt(btn.dataset.posteIdx);
        const jIdx = parseInt(btn.dataset.dayIdx);
        this._prefillDayFromEstim(pIdx, jIdx);
      });
    });

    document.getElementById('realiseMontantFacture')?.addEventListener('input', () => this._recompute());
    document.getElementById('realiseBody')?.addEventListener('input', () => this._recompute());
    document.getElementById('realiseBody')?.addEventListener('change', () => this._recompute());
  },

  /** Construit la grille de chips cochables pour un type (personnel/vehicules/materiel) */
  _buildRealiseChips(grid, type, preEntries) {
    grid.innerHTML = '';
    const defaults = CONFIG.LISTS[type] || [];
    const customs = CustomItems.load()[type] || [];
    const tarifs = (TarifManager.getData()?.items?.[type] || []).map(t => t.item).filter(Boolean);
    // Pour vehicules, on inclut aussi les engins pour faciliter
    let allItems = [...defaults, ...customs, ...tarifs];
    if (type === 'vehicules') {
      const enginsList = (TarifManager.getData()?.items?.engins || []).map(t => t.item).filter(Boolean);
      allItems = [...allItems, ...enginsList];
    }
    const seen = new Set();
    const items = [];
    allItems.forEach(it => {
      const k = (it || '').trim().toLowerCase();
      if (k && !seen.has(k)) { seen.add(k); items.push(it); }
    });

    // Map des items pré-cochés : nom → qty
    const preMap = {};
    preEntries.forEach(e => {
      const m = String(e).match(/^(\d+)x?\s+(.+?)(?:\s+\[.+?\])?$/i);
      if (m) preMap[m[2].trim().toLowerCase()] = parseInt(m[1]) || 1;
    });

    items.forEach(item => {
      const k = item.toLowerCase();
      const checked = preMap[k] !== undefined;
      const qty = preMap[k] || 1;
      const chip = document.createElement('label');
      chip.className = 'rchip' + (checked ? ' active' : '');
      chip.innerHTML = `
        <input type="checkbox" value="${this._esc(item)}" ${checked ? 'checked' : ''}>
        <span class="rchip-name">${this._esc(item)}</span>
        <input type="number" class="rchip-qty ${checked ? '' : 'hidden'}" value="${qty}" min="1">
      `;
      const cb = chip.querySelector('input[type="checkbox"]');
      const qtyInput = chip.querySelector('.rchip-qty');
      cb.addEventListener('change', () => {
        if (cb.checked) {
          chip.classList.add('active');
          qtyInput.classList.remove('hidden');
          qtyInput.focus(); qtyInput.select();
        } else {
          chip.classList.remove('active');
          qtyInput.classList.add('hidden');
          qtyInput.value = '1';
        }
      });
      grid.appendChild(chip);
    });
  },

  /** Pré-remplit un jour avec l'estimé du poste */
  _prefillDayFromEstim(pIdx, jIdx) {
    const postes = PosteManager.collectAll().filter(p => p.mode === 'detail');
    const poste = postes[pIdx];
    if (!poste) return;
    ['personnel', 'vehicules', 'materiel'].forEach(type => {
      const grid = document.querySelector(`.realise-chips[data-rtype="${type}"][data-poste-idx="${pIdx}"][data-day-idx="${jIdx}"]`);
      if (!grid) return;
      let estItems = [];
      if (type === 'personnel') estItems = poste.personnel || [];
      else if (type === 'vehicules') estItems = [...(poste.vehicules || []), ...(poste.engins || [])];
      else if (type === 'materiel') estItems = poste.materiel || [];
      this._buildRealiseChips(grid, type, estItems);
    });
    this._recompute();
    Toast.success('Jour pré-rempli depuis l\'estimé');
  },

  /** Lit les chips cochés d'une grille → array ["2x Item", ...] */
  _collectChips(grid) {
    if (!grid) return [];
    const out = [];
    grid.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
      const qty = cb.closest('.rchip')?.querySelector('.rchip-qty')?.value || '1';
      out.push(`${qty}x ${cb.value}`);
    });
    return out;
  },

  /** Génère la liste des jours d'un poste à partir des RDVs et nbJours */
  _expandJours(poste, existingJours) {
    const nbJours = parseInt(poste.jours) || 1;
    const rdvs = (poste.rdvs || []).filter(r => r.date);
    const dates = [];
    if (rdvs.length > 0) {
      const start = new Date(rdvs[0].date);
      // Génère N jours ouvrés
      const days = this._businessDays(start, nbJours);
      days.forEach(d => dates.push(this._iso(d)));
    } else {
      for (let i = 0; i < nbJours; i++) dates.push('');
    }
    return dates.map((date, i) => {
      const ex = existingJours[i] || {};
      return {
        date,
        personnel: ex.personnel || '',
        vehicules: ex.vehicules || '',
        materiel: ex.materiel || ''
      };
    });
  },

  _businessDays(start, n) {
    const out = [];
    const cur = new Date(start.getFullYear(), start.getMonth(), start.getDate());
    while (out.length < n) {
      const dow = cur.getDay();
      if (dow !== 0 && dow !== 6) out.push(new Date(cur));
      cur.setDate(cur.getDate() + 1);
    }
    return out;
  },

  _iso(d) {
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${d.getFullYear()}-${m}-${dd}`;
  },

  /** Parse "2x Manutentionnaire, 1x Chauffeur" → array d'items */
  _parseEntries(str) {
    if (!str) return [];
    return String(str).split(/[,;]+/).map(s => s.trim()).filter(s => s);
  },

  /** Calcule le coût d'une ligne réelle pour un type donné */
  _coutLigne(type, str) {
    let total = 0;
    this._parseEntries(str).forEach(entry => {
      const m = entry.match(/^(\d+)x?\s+(.+)$/i);
      if (!m) return;
      const qty = parseInt(m[1]) || 1;
      const nom = m[2].trim();
      // 1 jour par défaut pour les ressources réelles
      const cout = TarifManager.getCout(type, nom);
      total += qty * cout;
    });
    return total;
  },

  _recompute() {
    // Coût estimé = somme des coûts initiaux des postes
    const postes = PosteManager.collectAll().filter(p => p.mode === 'detail');
    let coutEstim = 0;
    postes.forEach((p, pIdx) => {
      const card = document.querySelectorAll('.poste-card')[pIdx];
      if (card) {
        const calc = TarifManager.calculerPrixPoste(card);
        coutEstim += calc.cout;
      }
    });

    // Coût réel = somme des coûts saisis dans les chips jour par jour
    let coutReel = 0;
    document.querySelectorAll('.realise-day').forEach(day => {
      const personnelEntries = this._collectChips(day.querySelector('.realise-chips[data-rtype="personnel"]'));
      const vehiculesEntries = this._collectChips(day.querySelector('.realise-chips[data-rtype="vehicules"]'));
      const materielEntries = this._collectChips(day.querySelector('.realise-chips[data-rtype="materiel"]'));

      personnelEntries.forEach(e => {
        const m = e.match(/^(\d+)x?\s+(.+)$/i);
        if (!m) return;
        coutReel += (parseInt(m[1]) || 1) * TarifManager.getCout('personnel', m[2].trim());
      });
      vehiculesEntries.forEach(e => {
        const m = e.match(/^(\d+)x?\s+(.+)$/i);
        if (!m) return;
        const qty = parseInt(m[1]) || 1;
        const nom = m[2].trim();
        coutReel += qty * (TarifManager.getCout('vehicules', nom) || TarifManager.getCout('engins', nom));
      });
      materielEntries.forEach(e => {
        const m = e.match(/^(\d+)x?\s+(.+)$/i);
        if (!m) return;
        coutReel += (parseInt(m[1]) || 1) * TarifManager.getCout('materiel', m[2].trim());
      });
    });

    const fmt = (v) => v.toLocaleString('fr-CH', { minimumFractionDigits: 0, maximumFractionDigits: 0 }) + ' CHF';
    document.getElementById('rsCoutEstim').textContent = fmt(coutEstim);
    document.getElementById('rsCoutReel').textContent = fmt(coutReel);

    // Delta coût
    const deltaCout = coutReel - coutEstim;
    const deltaCoutPct = coutEstim > 0 ? ((deltaCout / coutEstim) * 100).toFixed(1) : 0;
    const deltaEl = document.getElementById('rsCoutDelta');
    deltaEl.textContent = (deltaCout >= 0 ? '+' : '') + fmt(deltaCout) + ' (' + deltaCoutPct + '%)';
    deltaEl.className = 'rs-delta ' + (deltaCout > 0 ? 'negative' : 'positive');

    // Marge réelle = (montantFacture - coutReel) / montantFacture
    const montantFact = parseFloat(document.getElementById('realiseMontantFacture')?.value) || 0;
    const margeEl = document.getElementById('rsMargeReel');
    const margeDeltaEl = document.getElementById('rsMargeDelta');
    if (montantFact > 0 && coutReel > 0) {
      const margePct = ((montantFact - coutReel) / montantFact * 100).toFixed(1);
      const margeMontant = montantFact - coutReel;
      margeEl.textContent = margePct + ' %';
      margeDeltaEl.textContent = (margeMontant >= 0 ? '+' : '') + fmt(margeMontant);
      margeDeltaEl.className = 'rs-delta ' + (margeMontant > 0 ? 'positive' : 'negative');
    } else {
      margeEl.textContent = '—';
      margeDeltaEl.textContent = '';
    }
  },

  _fillFromEstim() {
    const ht = parseFloat(document.querySelector('[name="montantHT"]')?.value) || 0;
    document.getElementById('realiseMontantFacture').value = ht.toFixed(2);
    this._recompute();
  },

  _fillFromReel() {
    // Coût réel × marge moyenne (ex: 1.4)
    const txt = document.getElementById('rsCoutReel').textContent;
    const coutReel = parseFloat(txt.replace(/[^\d.]/g, '')) || 0;
    const margeMoyenne = 1.4; // par défaut +40%
    document.getElementById('realiseMontantFacture').value = (coutReel * margeMoyenne).toFixed(2);
    this._recompute();
  },

  _esc(s) {
    const div = document.createElement('div');
    div.textContent = String(s || '');
    return div.innerHTML;
  },

  /** Collecte les données saisies dans le modal (depuis les chips) */
  _collect() {
    const postes = [];
    document.querySelectorAll('.realise-poste').forEach(pEl => {
      const jours = [];
      pEl.querySelectorAll('.realise-day').forEach(dEl => {
        const personnelArr = this._collectChips(dEl.querySelector('.realise-chips[data-rtype="personnel"]'));
        const vehiculesArr = this._collectChips(dEl.querySelector('.realise-chips[data-rtype="vehicules"]'));
        const materielArr = this._collectChips(dEl.querySelector('.realise-chips[data-rtype="materiel"]'));
        jours.push({
          date: dEl.querySelector('.realise-day-date')?.textContent?.trim() || '',
          // Format string pour compatibilité descendante avec _coutLigne / Excel
          personnel: personnelArr.join(', '),
          vehicules: vehiculesArr.join(', '),
          materiel: materielArr.join(', '),
          notes: dEl.querySelector('.rj-notes')?.value || ''
        });
      });
      postes.push({ jours });
    });
    return {
      postes,
      montantFacture: parseFloat(document.getElementById('realiseMontantFacture').value) || 0,
      notes: document.getElementById('realiseNotes').value || '',
      savedAt: new Date().toISOString()
    };
  },

  async save() {
    const ref = document.querySelector('[name="ref"]')?.value?.trim();
    if (!ref) { Toast.error('Référence dossier requise.'); return; }
    const realise = this._collect();

    const fd = new FormData(document.getElementById('devisForm'));
    const data = Object.fromEntries(fd.entries());
    data.postes = PosteManager.collectAll();
    data.userId = UserManager.getUserId();
    data.realise = realise;
    data._action = 'save_realise';

    const btn = document.getElementById('realiseSave');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Sauvegarde…'; }

    try {
      const resp = await fetch(CONFIG.SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify(data)
      });
      const res = await resp.json();
      if (res.status === 'success') {
        Toast.success(`📋 Réalisé sauvegardé · marge ${this._currentMarge()}`);
        window._currentDossierRealise = realise;

        // Sauvegarde locale (localStorage) en backup
        try {
          const key = 'pelichet_realise_' + ref;
          localStorage.setItem(key, JSON.stringify({ ref, realise, savedAt: realise.savedAt }));
        } catch (e) { /* quota plein, ignore */ }

        // Téléchargement automatique du XLSX
        if (res.xlsxB64 && res.xlsxName) {
          FormSubmitter._downloadBase64File(
            res.xlsxB64,
            res.xlsxName,
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
          );
        }

        this.close();
        if (typeof DossierList !== 'undefined') {
          DossierList._loaded = false;
          DossierList.fetch();
        }
      } else {
        Toast.error('Erreur : ' + (res.message || 'Sauvegarde échouée'));
      }
    } catch (err) {
      Toast.error('Erreur de connexion : ' + err.message);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'Sauvegarder le réalisé'; }
    }
  },

  /** Restaure le réalisé depuis localStorage si pas encore en mémoire */
  _restoreFromLocal(ref) {
    try {
      const key = 'pelichet_realise_' + ref;
      const raw = localStorage.getItem(key);
      if (raw) return JSON.parse(raw).realise;
    } catch (e) { /* ignore */ }
    return null;
  },

  _currentMarge() {
    return document.getElementById('rsMargeReel')?.textContent || '—';
  }
};

// ============================================
// API KEYS MANAGER (settings panel)
// ============================================
const ApiKeysManager = {
  init() {
    // Onglets settings
    document.querySelectorAll('.settings-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        document.querySelectorAll('.settings-tab').forEach(t => t.classList.remove('active'));
        document.querySelectorAll('.settings-pane').forEach(p => p.classList.remove('active'));
        tab.classList.add('active');
        document.querySelector(`.settings-pane[data-spane="${tab.dataset.stab}"]`)?.classList.add('active');
        if (tab.dataset.stab === 'apikeys') this.refresh();
      });
    });

    document.getElementById('apikeyCreateBtn')?.addEventListener('click', () => this.createKey());

    // Remplir l'URL d'exemple dans la doc
    document.querySelectorAll('#apiUrlExample, .apiUrlExample').forEach(el => {
      el.textContent = CONFIG.SCRIPT_URL;
    });
  },

  async refresh() {
    const tbody = document.getElementById('apikeyTbody');
    if (!tbody) return;
    tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;color:var(--ink-3);padding:20px">Chargement…</td></tr>';
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=apikey_list`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status !== 'success') {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--ink-3);padding:20px">Erreur : ${result.message || 'inconnue'}</td></tr>`;
        return;
      }
      const keys = result.data || [];
      if (keys.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--ink-3);padding:20px">Aucune clé. Crée-en une ci-dessus.</td></tr>`;
        return;
      }
      const fmtDate = (s) => s ? new Date(s).toLocaleDateString('fr-CH', { day: '2-digit', month: 'short', year: 'numeric' }) : '—';
      tbody.innerHTML = keys.map(k => `
        <tr class="${k.active ? '' : 'revoked'}" data-key="${k.key}">
          <td class="name-cell">${this._esc(k.name)}</td>
          <td class="key-cell" title="${k.key}">${k.keyMasked}</td>
          <td><span class="apikey-pill perm-${k.permission}">${k.permission}</span></td>
          <td>${fmtDate(k.createdAt)}</td>
          <td>${fmtDate(k.lastUsed)}</td>
          <td>${k.active ? '<span class="status-active">Active</span>' : '<span class="status-revoked">Révoquée</span>'}</td>
          <td class="actions">
            <button type="button" data-act="copy" title="Copier la clé">📋</button>
            ${k.active
              ? '<button type="button" data-act="revoke" title="Révoquer">🚫</button>'
              : '<button type="button" data-act="reactivate" title="Réactiver">↻</button>'}
            <button type="button" data-act="delete" class="danger" title="Supprimer">🗑️</button>
          </td>
        </tr>
      `).join('');
      tbody.querySelectorAll('button[data-act]').forEach(btn => {
        btn.addEventListener('click', (e) => {
          e.stopPropagation();
          const tr = btn.closest('tr');
          this._action(btn.dataset.act, tr.dataset.key, keys.find(kk => kk.key === tr.dataset.key));
        });
      });
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="7" style="text-align:center;color:var(--ink-3);padding:20px">Erreur connexion : ${err.message}</td></tr>`;
    }
  },

  async createKey() {
    const name = document.getElementById('apikeyNewName')?.value.trim();
    const perm = document.getElementById('apikeyNewPerm')?.value;
    if (!name) { Toast.warning('Nom requis'); return; }
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=apikey_create&name=${encodeURIComponent(name)}&permission=${perm}`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && result.data) {
        const result_div = document.getElementById('apikeyNewResult');
        result_div.innerHTML = `
          <strong>✅ Clé créée pour "${this._esc(result.data.name)}"</strong>
          <div class="key-display">
            <code id="newKeyValue">${result.data.key}</code>
            <button type="button" class="btn btn-primary btn-xs" id="copyNewKey">📋 Copier</button>
          </div>
          <div class="warn">⚠️ Cette clé ne sera plus affichée en clair après cette page. Copie-la maintenant.</div>
        `;
        result_div.classList.remove('hidden');
        document.getElementById('copyNewKey')?.addEventListener('click', () => {
          navigator.clipboard?.writeText(result.data.key).then(() => Toast.success('Clé copiée'));
        });
        document.getElementById('apikeyNewName').value = '';
        this.refresh();
      } else {
        Toast.error(result.message || 'Erreur création');
      }
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    }
  },

  async _action(act, key, info) {
    if (act === 'copy') {
      navigator.clipboard?.writeText(key).then(() => Toast.success('Clé copiée'));
      return;
    }
    if (act === 'revoke') {
      if (!confirm(`Révoquer la clé "${info.name}" ? L'application qui l'utilise ne pourra plus accéder à l'API.`)) return;
      const resp = await fetch(`${CONFIG.SCRIPT_URL}?action=apikey_revoke&key=${encodeURIComponent(key)}`);
      const r = await resp.json();
      if (r.status === 'success') { Toast.success('Clé révoquée'); this.refresh(); }
      else Toast.error(r.message);
    }
    if (act === 'reactivate') {
      const resp = await fetch(`${CONFIG.SCRIPT_URL}?action=apikey_reactivate&key=${encodeURIComponent(key)}`);
      const r = await resp.json();
      if (r.status === 'success') { Toast.success('Clé réactivée'); this.refresh(); }
      else Toast.error(r.message);
    }
    if (act === 'delete') {
      if (!confirm(`Supprimer DÉFINITIVEMENT la clé "${info.name}" ?`)) return;
      const resp = await fetch(`${CONFIG.SCRIPT_URL}?action=apikey_delete&key=${encodeURIComponent(key)}`);
      const r = await resp.json();
      if (r.status === 'success') { Toast.success('Clé supprimée'); this.refresh(); }
      else Toast.error(r.message);
    }
  },

  _esc(s) {
    const d = document.createElement('div');
    d.textContent = String(s || '');
    return d.innerHTML;
  }
};

// ============================================
// AFFAIRES VIEW (tableau d'affaires + forecast)
// ============================================
const AffairesView = {
  _affaires: [],
  _sortKey: 'datePrevue',
  _sortDir: 'desc',
  _forecastStart: null, // 'YYYY-MM'
  _forecastEnd: null,   // 'YYYY-MM'

  init() {
    document.getElementById('affairesBtn')?.addEventListener('click', () => this.open());
    document.getElementById('affairesClose')?.addEventListener('click', () => this.close());
    document.getElementById('affairesSearch')?.addEventListener('input', () => this._render());
    document.getElementById('affairesStatut')?.addEventListener('change', () => this._render());
    document.getElementById('affairesExportCsv')?.addEventListener('click', () => this._exportCsv());
    document.querySelectorAll('.affaires-table th.sortable').forEach(th => {
      th.addEventListener('click', () => {
        const key = th.dataset.sort;
        if (this._sortKey === key) {
          this._sortDir = this._sortDir === 'asc' ? 'desc' : 'asc';
        } else {
          this._sortKey = key;
          this._sortDir = key.includes('mont') || key.includes('cout') || key.includes('gain') || key.includes('marge') ? 'desc' : 'asc';
        }
        this._render();
      });
    });

    // Plage personnalisée du forecast
    document.getElementById('affairesForecastStart')?.addEventListener('change', (e) => {
      this._forecastStart = e.target.value || null;
      this._clearPreset();
      this._render();
    });
    document.getElementById('affairesForecastEnd')?.addEventListener('change', (e) => {
      this._forecastEnd = e.target.value || null;
      this._clearPreset();
      this._render();
    });
    document.getElementById('affairesForecastReset')?.addEventListener('click', () => {
      this._applyPreset('12');
    });
    document.querySelectorAll('.forecast-presets .btn').forEach(btn => {
      btn.addEventListener('click', () => this._applyPreset(btn.dataset.preset));
    });
  },

  _clearPreset() {
    document.querySelectorAll('.forecast-presets .btn').forEach(b => b.classList.remove('active'));
  },

  _applyPreset(preset) {
    const today = new Date();
    today.setDate(1);
    const fmtMonth = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;

    if (preset === '3' || preset === '6' || preset === '12') {
      const n = parseInt(preset);
      this._forecastStart = fmtMonth(today);
      const end = new Date(today.getFullYear(), today.getMonth() + n - 1, 1);
      this._forecastEnd = fmtMonth(end);
    } else if (preset === 'ytd') {
      this._forecastStart = `${today.getFullYear()}-01`;
      this._forecastEnd = `${today.getFullYear()}-12`;
    } else if (preset === 'all') {
      // Utiliser la plage min-max des affaires
      const months = this._affaires.map(a => a.mois).filter(Boolean).sort();
      this._forecastStart = months[0] || fmtMonth(today);
      this._forecastEnd = months[months.length - 1] || fmtMonth(today);
    }

    document.getElementById('affairesForecastStart').value = this._forecastStart || '';
    document.getElementById('affairesForecastEnd').value = this._forecastEnd || '';

    this._clearPreset();
    document.querySelector(`.forecast-presets .btn[data-preset="${preset}"]`)?.classList.add('active');
    this._render();
  },

  async open() {
    document.getElementById('affairesPanel')?.classList.remove('hidden');
    document.getElementById('affairesTbody').innerHTML = '<tr><td colspan="8" class="affaires-empty">Chargement…</td></tr>';

    // Plage forecast par défaut : 12 mois à venir
    if (!this._forecastStart) {
      const today = new Date();
      today.setDate(1);
      const fmtMonth = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      this._forecastStart = fmtMonth(today);
      const end = new Date(today.getFullYear(), today.getMonth() + 11, 1);
      this._forecastEnd = fmtMonth(end);
      document.getElementById('affairesForecastStart').value = this._forecastStart;
      document.getElementById('affairesForecastEnd').value = this._forecastEnd;
    }

    try {
      const url = `${CONFIG.SCRIPT_URL}?action=affaires_list&user=${encodeURIComponent(UserManager.getUserId())}`;
      const resp = await fetch(url);
      const result = await resp.json();
      if (result.status === 'success' && Array.isArray(result.data)) {
        this._affaires = result.data;
        this._render();
      } else {
        document.getElementById('affairesTbody').innerHTML = `<tr><td colspan="8" class="affaires-empty">Erreur : ${result.message || 'inconnue'}</td></tr>`;
      }
    } catch (err) {
      document.getElementById('affairesTbody').innerHTML = `<tr><td colspan="8" class="affaires-empty">Erreur de connexion : ${err.message}</td></tr>`;
    }
  },

  close() { document.getElementById('affairesPanel')?.classList.add('hidden'); },

  _filtered() {
    const q = (document.getElementById('affairesSearch')?.value || '').toLowerCase().trim();
    const statutFilter = (document.getElementById('affairesStatut')?.value || '').trim();
    return this._affaires.filter(a => {
      if (q && !((a.ref || '').toLowerCase().includes(q) || (a.client || '').toLowerCase().includes(q))) return false;
      if (statutFilter) {
        const s = (a.statut || '').toLowerCase();
        const f = statutFilter.toLowerCase();
        if (f === 'refusé' && !s.includes('refus') && !s.includes('annul')) return false;
        else if (f === 'ok' && !s.includes('ok') && !s.includes('accept')) return false;
        else if (f === 'réalisé' && !s.includes('réalis') && !s.includes('realis')) return false;
        else if (f === 'brouillon' && !s.includes('brouillon')) return false;
      }
      return true;
    });
  },

  _sort(arr) {
    const key = this._sortKey;
    const dir = this._sortDir === 'asc' ? 1 : -1;
    return arr.slice().sort((a, b) => {
      let va = a[key], vb = b[key];
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va == null) va = '';
      if (vb == null) vb = '';
      if (va < vb) return -1 * dir;
      if (va > vb) return 1 * dir;
      return 0;
    });
  },

  _statutClass(s) {
    const lo = (s || '').toLowerCase();
    if (lo.includes('refus') || lo.includes('annul')) return 'statut-refus';
    if (lo.includes('réalis') || lo.includes('realis')) return 'statut-realise';
    if (lo.includes('brouillon')) return 'statut-brouillon';
    if (lo.includes('ok') || lo.includes('accept')) return 'statut-ok';
    return 'statut-brouillon';
  },

  _render() {
    const fmt = (v) => Math.round(v || 0).toLocaleString('fr-CH');
    const fmtPct = (v) => (v >= 0 ? '+' : '') + (v || 0).toFixed(1) + ' %';
    const fmtDate = (d) => {
      if (!d) return '—';
      const m = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
      return m ? `${m[3]}.${m[2]}.${m[1]}` : d;
    };

    // Filtre + tri
    const filtered = this._sort(this._filtered());

    // KPIs (sur les filtrés)
    const totalCA = filtered.reduce((s, a) => s + (a.montantFacture || 0), 0);
    const totalCout = filtered.reduce((s, a) => s + (a.coutPourMarge || 0), 0);
    const totalGain = totalCA - totalCout;
    const margeGlobale = totalCA > 0 ? (totalGain / totalCA * 100) : 0;
    const nbAffaires = filtered.length;
    const nbAcceptes = filtered.filter(a => /ok|accept/i.test(a.statut)).length;
    const nbRefuses = filtered.filter(a => /refus|annul/i.test(a.statut)).length;

    const kpiHtml = `
      <div class="kpi-card">
        <div class="kpi-label">Chiffre d'affaires</div>
        <div class="kpi-value"><span class="cur">CHF</span>${fmt(totalCA)}</div>
        <div class="kpi-sub">${nbAffaires} affaire${nbAffaires > 1 ? 's' : ''}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Coût total</div>
        <div class="kpi-value"><span class="cur">CHF</span>${fmt(totalCout)}</div>
        <div class="kpi-sub">${filtered.filter(a => a.hasRealise).length} avec réalisé</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Gain</div>
        <div class="kpi-value ${totalGain >= 0 ? '' : 'pct negative'}"><span class="cur">CHF</span>${fmt(totalGain)}</div>
        <div class="kpi-sub">${totalGain >= 0 ? 'Bénéfice' : 'Perte'}</div>
      </div>
      <div class="kpi-card">
        <div class="kpi-label">Marge moyenne</div>
        <div class="kpi-value pct ${margeGlobale < 0 ? 'negative' : ''}">${fmtPct(margeGlobale)}</div>
        <div class="kpi-sub">${nbAcceptes} OK · ${nbRefuses} refusés</div>
      </div>
    `;
    document.getElementById('affairesKpis').innerHTML = kpiHtml;

    // Forecast par mois selon la plage choisie par l'utilisateur
    const today = new Date();
    today.setDate(1);
    const currentMonthKey = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;

    // Parse plage début/fin
    const parseMonth = (s) => {
      if (!s) return null;
      const m = s.match(/^(\d{4})-(\d{2})/);
      return m ? new Date(parseInt(m[1]), parseInt(m[2]) - 1, 1) : null;
    };
    let dStart = parseMonth(this._forecastStart);
    let dEnd = parseMonth(this._forecastEnd);
    if (!dStart) dStart = new Date(today);
    if (!dEnd) dEnd = new Date(today.getFullYear(), today.getMonth() + 11, 1);
    if (dEnd < dStart) { const tmp = dStart; dStart = dEnd; dEnd = tmp; }

    // Limite raisonnable : 60 mois max
    const totalMonths = (dEnd.getFullYear() - dStart.getFullYear()) * 12 + (dEnd.getMonth() - dStart.getMonth()) + 1;
    const nbMonths = Math.min(60, Math.max(1, totalMonths));

    const months = [];
    for (let i = 0; i < nbMonths; i++) {
      const d = new Date(dStart.getFullYear(), dStart.getMonth() + i, 1);
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      months.push({
        key: `${d.getFullYear()}-${mm}`,
        label: d.toLocaleDateString('fr-CH', { month: 'short' }) + ' ' + String(d.getFullYear()).slice(2),
        total: 0,
        realise: 0,
        nbAffaires: 0
      });
    }

    let totalForecast = 0;
    let totalRealise = 0;
    let nbInRange = 0;
    filtered.forEach(a => {
      if (!a.mois) return;
      const m = months.find(mo => mo.key === a.mois);
      if (!m) return;
      m.total += a.montantFacture || 0;
      m.nbAffaires++;
      totalForecast += a.montantFacture || 0;
      nbInRange++;
      if (a.hasRealise) {
        m.realise += a.montantFacture || 0;
        totalRealise += a.montantFacture || 0;
      }
    });

    const maxMonth = Math.max(1, ...months.map(m => m.total));
    const fc = document.getElementById('affairesForecast');
    fc.style.gridTemplateColumns = `repeat(${nbMonths}, minmax(60px, 1fr))`;
    // Format compact pour les montants ('45k', '1.2M')
    const fmtCompact = (v) => {
      v = Math.round(v || 0);
      if (v === 0) return '—';
      if (v >= 1000000) return (v / 1000000).toFixed(1) + 'M';
      if (v >= 1000) return Math.round(v / 1000) + 'k';
      return v.toString();
    };

    fc.innerHTML = months.map(m => {
      const pct = (m.total / maxMonth) * 100;
      const isCurrent = m.key === currentMonthKey;
      const hasRealise = m.realise > 0;
      const tooltip = `${m.label}: ${fmt(m.total)} CHF${m.nbAffaires > 0 ? ' (' + m.nbAffaires + ' affaire' + (m.nbAffaires > 1 ? 's' : '') + ')' : ''}`;
      const isZero = (m.total || 0) === 0;
      return `
        <div class="fc-month ${isCurrent ? 'current' : ''}" title="${tooltip}">
          <div class="fc-amount ${isZero ? 'zero' : ''}">${fmtCompact(m.total)}</div>
          <div class="fc-bar-wrap">
            <div class="fc-bar ${hasRealise ? 'has-realise' : ''}" style="height:${pct}%"></div>
          </div>
          <div class="fc-label">${m.label}</div>
        </div>
      `;
    }).join('');

    // Synthèse forecast
    const summaryEl = document.getElementById('affairesForecastSummary');
    if (summaryEl) {
      const startLabel = dStart.toLocaleDateString('fr-CH', { month: 'short', year: 'numeric' });
      const endLabel = dEnd.toLocaleDateString('fr-CH', { month: 'short', year: 'numeric' });
      summaryEl.innerHTML = `
        <span>Période :<strong>${startLabel} → ${endLabel}</strong></span>
        <span>·</span>
        <span>Affaires :<strong>${nbInRange}</strong></span>
        <span>·</span>
        <span>CA prévu :<strong>${fmt(totalForecast)} CHF</strong></span>
        <span>·</span>
        <span>CA réalisé :<strong>${fmt(totalRealise)} CHF</strong></span>
      `;
    }

    // Update header sort indicators
    document.querySelectorAll('.affaires-table th.sortable').forEach(th => {
      th.classList.remove('sorted-asc', 'sorted-desc');
      if (th.dataset.sort === this._sortKey) {
        th.classList.add(this._sortDir === 'asc' ? 'sorted-asc' : 'sorted-desc');
      }
    });

    // Table
    const tbody = document.getElementById('affairesTbody');
    if (filtered.length === 0) {
      tbody.innerHTML = '<tr><td colspan="8" class="affaires-empty">Aucune affaire trouvée</td></tr>';
      return;
    }
    tbody.innerHTML = filtered.map(a => `
      <tr data-ref="${a.ref}">
        <td class="ref-cell">${a.ref}</td>
        <td>${a.client || '—'}</td>
        <td><span class="statut-pill ${this._statutClass(a.statut)}">${a.statut || '—'}</span></td>
        <td>${fmtDate(a.datePrevue)}</td>
        <td class="num">${fmt(a.coutPourMarge)} CHF</td>
        <td class="num">${fmt(a.montantFacture)} CHF</td>
        <td class="num ${a.gain >= 0 ? 'gain-pos' : 'gain-neg'}">${a.gain >= 0 ? '+' : ''}${fmt(a.gain)} CHF</td>
        <td class="num ${a.margePct >= 0 ? 'marge-pos' : 'marge-neg'}">${fmtPct(a.margePct)}</td>
      </tr>
    `).join('');

    tbody.querySelectorAll('tr').forEach(tr => {
      tr.addEventListener('click', () => {
        const ref = tr.dataset.ref;
        if (ref) {
          this.close();
          DossierLoader.load(ref);
        }
      });
    });
  },

  _exportCsv() {
    const filtered = this._sort(this._filtered());
    const BOM = '﻿';
    const headers = ['Référence', 'Client', 'Statut', 'Date prévue', 'Coût total', 'Montant facturé', 'Gain', 'Marge %', 'Type société'];
    let csv = BOM + headers.join(';') + '\n';
    filtered.forEach(a => {
      csv += [
        a.ref,
        '"' + (a.client || '').replace(/"/g, '""') + '"',
        a.statut || '',
        a.datePrevue || '',
        Math.round(a.coutPourMarge || 0).toString().replace('.', ','),
        Math.round(a.montantFacture || 0).toString().replace('.', ','),
        Math.round(a.gain || 0).toString().replace('.', ','),
        (a.margePct || 0).toFixed(1).replace('.', ','),
        a.typeSociete || ''
      ].join(';') + '\n';
    });
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `affaires_${UserManager.getUserId()}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
    Toast.success('Export CSV téléchargé');
  }
};

// ============================================
// MOBILE SHELL (drawer sidebar + bottom bar)
// ============================================
const MobileShell = {
  init() {
    const menuBtn = document.getElementById('mobileMenuBtn');
    const backdrop = document.getElementById('mobileBackdrop');
    const sidebar = document.querySelector('.sidebar');

    if (menuBtn && backdrop && sidebar) {
      menuBtn.addEventListener('click', () => {
        sidebar.classList.toggle('open');
        backdrop.classList.toggle('hidden');
      });
      backdrop.addEventListener('click', () => {
        sidebar.classList.remove('open');
        backdrop.classList.add('hidden');
        this._collapseSearch();
      });

      // Fermer le drawer quand on clique sur un dossier
      document.getElementById('dossierListBody')?.addEventListener('click', (e) => {
        if (e.target.closest('.dossier-item')) {
          sidebar.classList.remove('open');
          backdrop.classList.add('hidden');
        }
      });
    }

    // Recherche mobile : tap icône loupe → expand la barre
    const searchWrap = document.querySelector('.topbar-search');
    const searchIcon = searchWrap?.querySelector('.search-icon');
    const searchInput = document.getElementById('searchRef');
    if (searchWrap && searchIcon) {
      searchIcon.addEventListener('click', (e) => {
        // Seulement en mobile (icône cliquable)
        if (window.innerWidth <= 768) {
          e.stopPropagation();
          searchWrap.classList.add('expanded');
          backdrop?.classList.remove('hidden');
          setTimeout(() => searchInput?.focus(), 50);
        }
      });
    }
    // Fermer la recherche quand on tape Escape ou clique ailleurs
    searchInput?.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') this._collapseSearch();
    });

    // Update bottom bar amount au rendu des postes / changement HT
    document.querySelector('[name="montantHT"]')?.addEventListener('input', () => this.updateBottomBar());
    document.addEventListener('input', (e) => {
      if (e.target?.name === 'postePrix' || e.target?.name === 'montantHT') this.updateBottomBar();
    });
    this.updateBottomBar();
  },

  updateBottomBar() {
    const el = document.getElementById('mbbAmount');
    if (!el) return;
    const ht = parseFloat(document.querySelector('[name="montantHT"]')?.value) || 0;
    const fmt = (v) => Math.round(v).toLocaleString('fr-CH');
    el.textContent = `CHF ${fmt(ht)}`;
  },

  _collapseSearch() {
    const wrap = document.querySelector('.topbar-search');
    const backdrop = document.getElementById('mobileBackdrop');
    const sidebar = document.querySelector('.sidebar');
    if (wrap?.classList.contains('expanded')) {
      wrap.classList.remove('expanded');
      // Cacher le backdrop uniquement si le sidebar n'est pas ouvert
      if (!sidebar?.classList.contains('open')) backdrop?.classList.add('hidden');
    }
  }
};

// ============================================
// PWA INSTALL PROMPT
// ============================================
const PWAInstall = {
  _deferredPrompt: null,

  init() {
    // Chrome / Android : beforeinstallprompt
    window.addEventListener('beforeinstallprompt', (e) => {
      e.preventDefault();
      this._deferredPrompt = e;
      this._showHint('android');
    });

    // iOS : pas d'API, afficher un hint manuel si mode non-standalone
    const isIOS = /iPhone|iPad|iPod/i.test(navigator.userAgent);
    const isStandalone = window.matchMedia('(display-mode: standalone)').matches
      || window.navigator.standalone === true;
    if (isIOS && !isStandalone && !localStorage.getItem('pelichet_hide_ios_install')) {
      setTimeout(() => this._showHint('ios'), 3000);
    }
  },

  _showHint(platform) {
    // Ne pas montrer plusieurs fois
    if (document.getElementById('installHint')) return;

    const hint = document.createElement('div');
    hint.id = 'installHint';
    hint.className = 'install-hint';
    if (platform === 'ios') {
      hint.innerHTML = `
        <span>📱 Installe Pelichet sur ton écran d'accueil — Partager ↑ puis "Sur l'écran d'accueil"</span>
        <button type="button" aria-label="Fermer">×</button>
      `;
      hint.querySelector('button').addEventListener('click', () => {
        hint.remove();
        localStorage.setItem('pelichet_hide_ios_install', '1');
      });
    } else {
      hint.innerHTML = `
        <span>📱 Installer l'application Pelichet ?</span>
        <button type="button" class="btn btn-red btn-xs">Installer</button>
        <button type="button" aria-label="Fermer">×</button>
      `;
      hint.querySelector('.btn-red').addEventListener('click', async () => {
        if (this._deferredPrompt) {
          this._deferredPrompt.prompt();
          await this._deferredPrompt.userChoice;
          this._deferredPrompt = null;
        }
        hint.remove();
      });
      hint.querySelectorAll('button[aria-label="Fermer"]').forEach(b =>
        b.addEventListener('click', () => hint.remove())
      );
    }
    document.body.appendChild(hint);
  }
};

// ============================================
// RIGHT RAIL (totaux + répartition + infos dossier)
// ============================================
const RightRail = {
  update() {
    const fmt = (v) => v.toLocaleString('fr-CH', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    const fmt0 = (v) => Math.round(v).toLocaleString('fr-CH');

    const htInput = document.querySelector('[name="montantHT"]');
    const ht = parseFloat(htInput?.value) || this._sumPostes();
    const tva = ht * (CONFIG.TVA_RATE || 0.081);
    const rplp = ht * (CONFIG.RPLP_RATE || 0.005);
    const ttc = ht + tva + rplp;

    // Calcul coûts par catégorie et marge (si TarifManager disponible)
    let coutTotal = 0;
    const catTotals = { personnel: 0, vehicules: 0, engins: 0, materiel: 0 };
    document.querySelectorAll('.poste-card').forEach(card => {
      try {
        const calc = TarifManager?.calculerPrixPoste?.(card);
        if (calc) {
          coutTotal += calc.cout || 0;
          (calc.details || []).forEach(d => {
            if (catTotals[d.type] !== undefined) catTotals[d.type] += d.prix || 0;
          });
        }
      } catch (e) { /* ignore */ }
    });

    const marge = ht > 0 && coutTotal > 0 ? ((ht - coutTotal) / ht * 100) : 0;
    const gain = ht - coutTotal;

    // Gros chiffre HT
    const totalHT = document.getElementById('railTotalHT');
    if (totalHT) totalHT.innerHTML = `<span class="cur">CHF</span>${fmt0(ht)}`;

    // Total sub
    const postesCount = document.querySelectorAll('.poste-card').length;
    const sub = document.getElementById('railTotalSub');
    if (sub) sub.textContent = `HT · ${postesCount} poste${postesCount > 1 ? 's' : ''} · ${fmt0(ttc)} TTC`;

    // Marge
    const margePct = document.getElementById('railMargePct');
    if (margePct) margePct.textContent = coutTotal > 0 ? `${marge.toFixed(1)} %` : '—';
    const margeBar = document.getElementById('railMargeBar');
    if (margeBar) margeBar.style.width = `${Math.max(0, Math.min(100, marge * 2))}%`;
    const railCout = document.getElementById('railCout');
    if (railCout) railCout.textContent = `Coût ${fmt(coutTotal)}`;
    const railGain = document.getElementById('railGain');
    if (railGain) railGain.textContent = `Gain ${fmt(gain)}`;

    // Répartition par catégorie
    const catBreak = document.getElementById('railCatBreak');
    if (catBreak) {
      const labels = { personnel: 'Personnel', vehicules: 'Véhicules', engins: 'Engins', materiel: 'Matériel' };
      catBreak.innerHTML = Object.keys(labels).map(k => {
        const v = catTotals[k] || 0;
        const pct = ht > 0 ? (v / ht * 100) : 0;
        return `
          <div class="cat-break">
            <div class="cat-break-row"><span class="k">${labels[k]}</span><span class="v">${fmt(v)}</span></div>
            <div class="cat-break-bar"><div class="cat-break-bar-fill" style="width:${pct}%"></div></div>
          </div>
        `;
      }).join('');
    }

    // Totaux HT/TVA/RPLP/TTC
    const setText = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
    setText('railHT', fmt(ht));
    setText('railTVA', fmt(tva));
    setText('railRPLP', fmt(rplp));
    setText('railTTC', fmt(ttc));

    // Infos dossier
    const vendeur = document.querySelector('[name="vendeur"]')?.value || '—';
    setText('railVendeur', vendeur);
    const volume = document.querySelector('[name="volumeEstime"]')?.value || '—';
    setText('railVolume', `${volume} m³`);
    const kmTxt = typeof KmCalculator !== 'undefined' && KmCalculator._lastResult
      ? `${KmCalculator._lastResult.totalKm} km` : '—';
    setText('railKm', kmTxt);
    setText('railPostesCount', String(postesCount));

    // Page head
    const client = document.querySelector('[name="client"]')?.value?.trim() || 'Nouveau devis';
    const ref = document.querySelector('[name="ref"]')?.value?.trim() || '—';
    setText('pageTitle', client);
    setText('pageRef', ref);
    setText('pagePostesCount', `${postesCount} poste${postesCount > 1 ? 's' : ''}`);

    // Bottom bar mobile
    setText('mbbAmount', `CHF ${Math.round(ht).toLocaleString('fr-CH')}`);
  },

  _sumPostes() {
    let total = 0;
    document.querySelectorAll('[name="postePrix"]').forEach(input => {
      const val = parseFloat(input.value);
      if (!isNaN(val)) total += val;
    });
    return total;
  },

  init() {
    // Reactiver sur changements de champs clés
    ['client', 'ref', 'volumeEstime'].forEach(name => {
      document.querySelector(`[name="${name}"]`)?.addEventListener('input', () => this.update());
    });
    // Actions boutons right rail
    document.getElementById('railExportVentil')?.addEventListener('click', () => VentilationExport.generate());
    document.getElementById('railPlanning')?.addEventListener('click', () => PlanningExport.generate());
    this.update();
  }
};

// ============================================
// POSTE MANAGER
// ============================================
const PosteManager = {
  _container: null,

  init() {
    this._container = document.getElementById('prestationsContainer');
    this.addPoste('detail');
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
        <select name="enginDuree" class="cb-duree" style="margin-left:0.5rem">
          <option value="">Durée mission</option>
          <option value="1/2 AM">½ AM</option>
          <option value="1/2 PM">½ PM</option>
          <option value="1j">1 jour</option>
          <option value="2j">2 jours</option>
          <option value="3j">3 jours</option>
          <option value="4j">4 jours</option>
          <option value="5j">5 jours</option>
        </select>
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

  /** Construit la grille de checkboxes pour un type (personnel, vehicules, materiel) */
  _buildCheckboxGrid(card, type) {
    const grid = card.querySelector(`.cb-grid[data-type="${type}"]`);
    if (!grid) return;
    grid.innerHTML = '';
    const defaults = CONFIG.LISTS[type] || [];
    const customs = (CustomItems.load()[type] || []);
    // Inclure aussi les items du panneau Tarifs (toute entrée tarif = item disponible)
    const tarifItems = (TarifManager.getData()?.items?.[type] || []).map(t => t.item).filter(Boolean);
    // Union sans doublons (préserve l'ordre : defaults puis customs puis tarifs)
    const allItems = [];
    const seen = new Set();
    [...defaults, ...customs, ...tarifItems].forEach(item => {
      const key = item.trim().toLowerCase();
      if (!seen.has(key) && item) { seen.add(key); allItems.push(item); }
    });
    allItems.forEach(item => this._addCheckboxItem(grid, item, false, 1, type));
  },

  /** Ajoute une checkbox item dans la grille */
  _addCheckboxItem(grid, itemName, checked = false, qty = 1, type = '', duree = '') {
    const hasDuree = (type === 'personnel' || type === 'vehicules');
    const div = document.createElement('div');
    div.className = 'cb-item' + (checked ? ' active' : '') + (hasDuree ? ' has-duree' : '');

    // Header : checkbox + nom
    const header = document.createElement('div');
    header.className = 'cb-header';
    header.innerHTML = `
      <input type="checkbox" value="${itemName}" ${checked ? 'checked' : ''}>
      <span class="cb-name">${itemName}</span>
    `;
    div.appendChild(header);

    if (hasDuree) {
      // Conteneur de lignes durée (chaque ligne = qté + durée + boutons)
      const rowsContainer = document.createElement('div');
      rowsContainer.className = 'cb-rows' + (checked ? '' : ' hidden');
      div.appendChild(rowsContainer);
      this._addDureeRow(rowsContainer, qty, duree, true);
    } else {
      // Matériel : juste un champ quantité simple
      const qtyInput = document.createElement('input');
      qtyInput.type = 'number';
      qtyInput.name = 'cbQty';
      qtyInput.value = qty;
      qtyInput.min = '1';
      qtyInput.className = 'cb-qty' + (checked ? '' : ' hidden');
      qtyInput.addEventListener('click', (e) => e.stopPropagation());
      div.appendChild(qtyInput);
    }

    const cb = div.querySelector('input[type="checkbox"]');

    // Clic sur le header = toggle checkbox
    header.addEventListener('click', (e) => {
      if (e.target === cb) return;
      cb.checked = !cb.checked;
      cb.dispatchEvent(new Event('change'));
    });

    cb.addEventListener('change', () => {
      if (cb.checked) {
        div.classList.add('active');
        if (hasDuree) {
          div.querySelector('.cb-rows')?.classList.remove('hidden');
        } else {
          const qi = div.querySelector('.cb-qty');
          if (qi) { qi.classList.remove('hidden'); qi.focus(); qi.select(); }
        }
      } else {
        div.classList.remove('active');
        if (hasDuree) {
          const rows = div.querySelector('.cb-rows');
          if (rows) { rows.classList.add('hidden'); rows.innerHTML = ''; PosteManager._addDureeRow(rows, 1, '', true); }
        } else {
          const qi = div.querySelector('.cb-qty');
          if (qi) { qi.classList.add('hidden'); qi.value = '1'; }
        }
      }
    });

    grid.appendChild(div);
    return div;
  },

  /** Génère le HTML d'un <select> durée */
  _dureeOptionsHTML(duree = '') {
    const opts = [
      ['', 'Durée mission'],
      ['1/2 AM', '½ AM'], ['1/2 PM', '½ PM'],
      ['1j', '1j'], ['2j', '2j'], ['3j', '3j'], ['4j', '4j'], ['5j', '5j']
    ];
    return opts.map(([v, t]) => `<option value="${v}" ${duree === v ? 'selected' : ''}>${t}</option>`).join('');
  },

  /** Ajoute une ligne qté+durée dans un conteneur .cb-rows */
  _addDureeRow(container, qty = 1, duree = '', isFirst = true) {
    const row = document.createElement('div');
    row.className = 'cb-row';
    row.innerHTML = `
      <input type="number" class="cb-qty" value="${qty}" min="1">
      <select class="cb-duree">${this._dureeOptionsHTML(duree)}</select>
      <button type="button" class="cb-row-btn add" title="Ajouter une ligne">+</button>
      ${isFirst ? '' : '<button type="button" class="cb-row-btn remove" title="Supprimer">−</button>'}
    `;
    row.querySelector('.cb-row-btn.add').addEventListener('click', (e) => {
      e.stopPropagation();
      PosteManager._addDureeRow(container, 1, '', false);
    });
    const removeBtn = row.querySelector('.cb-row-btn.remove');
    if (removeBtn) {
      removeBtn.addEventListener('click', (e) => { e.stopPropagation(); row.remove(); });
    }
    // Stop propagation sur tous les éléments interactifs pour éviter le toggle checkbox
    row.querySelectorAll('input, select, button').forEach(el => {
      el.addEventListener('click', (e) => e.stopPropagation());
      el.addEventListener('mousedown', (e) => e.stopPropagation());
    });
    container.appendChild(row);
    return row;
  },

  /** Ajouter un item personnalisé via prompt */
  _addCustomCbItem(card, type) {
    const labels = { personnel: 'personnel', vehicules: 'véhicule', materiel: 'matériel' };
    const name = prompt('Nom du ' + (labels[type] || type) + ' :');
    if (!name || !name.trim()) return;
    const val = name.trim();
    CustomItems.add(type, val);
    const grid = card.querySelector(`.cb-grid[data-type="${type}"]`);
    this._addCheckboxItem(grid, val, true, 1, type);
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

  addPoste(mode = 'detail') {
    const div = document.createElement('div');
    div.className = 'poste-card poste-detail';
    div.dataset.mode = 'detail';

    div.innerHTML = `
      <span class="poste-number"></span>
      <button type="button" class="poste-remove" title="Supprimer le poste">&times;</button>
      <div class="grid-2 mb-3">
        <div class="col-span-2" style="display:grid;grid-template-columns:1fr auto;gap:0.5rem">
          <input type="text" name="posteTitre" placeholder="DESIGNATION" class="font-bold uppercase">
          <input type="number" name="postePrix" placeholder="Prix HT" style="width:7rem;text-align:right" class="font-bold">
        </div>
      </div>
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
          <div class="cb-grid" data-type="personnel"></div>
          <button type="button" class="add-cb-custom" data-type="personnel">+ Autre personnel</button>
        </div>
        <div class="detail-section vehicules">
          <label>Véhicules :</label>
          <div class="cb-grid" data-type="vehicules"></div>
          <button type="button" class="add-cb-custom" data-type="vehicules">+ Autre véhicule</button>
        </div>
        <div class="detail-section materiel">
          <label>Matériel :</label>
          <div class="cb-grid" data-type="materiel"></div>
          <button type="button" class="add-cb-custom" data-type="materiel">+ Autre matériel</button>
        </div>
        <textarea name="tache" rows="2" placeholder="Instructions..." class="text-sm"></textarea>
        <div class="prix-auto-info hidden">
          <div class="prix-auto-detail"></div>
          <button type="button" class="prix-auto-apply">Appliquer le prix calculé</button>
        </div>
      </div>
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

    // Event bindings
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

    // Engins : dropdown (tonnage spécifique)
    const enginsSection = div.querySelector('.detail-section.engins');
    enginsSection.querySelector('.add-row-btn').addEventListener('click', () => {
      const row = this._createRow('engins');
      enginsSection.querySelector('.rows').appendChild(row);
      row.querySelector('select')?.addEventListener('change', recalc);
      row.querySelector('input[name="qty"]')?.addEventListener('input', recalc);
    });

    // Personnel, Véhicules, Matériel : grilles de checkboxes
    ['personnel', 'vehicules', 'materiel'].forEach(type => {
      this._buildCheckboxGrid(div, type);
      div.querySelector(`.add-cb-custom[data-type="${type}"]`).addEventListener('click', () => {
        this._addCustomCbItem(div, type);
      });
      // Recalculer le prix à chaque changement dans les grilles (checkbox, qté, durée)
      const grid = div.querySelector(`.cb-grid[data-type="${type}"]`);
      if (grid) {
        grid.addEventListener('change', recalc);
        grid.addEventListener('input', recalc);
      }
    });

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
    const hasMissing = calc.missing && calc.missing.length > 0;

    if (calc.prix === 0 && calc.cout === 0 && !hasMissing) {
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

    if (hasMissing) {
      const labels = { personnel: 'Personnel', vehicules: 'Véhicules', engins: 'Engins', materiel: 'Matériel' };
      const list = calc.missing.map(m => `<li><strong>${labels[m.type] || m.type}</strong> · ${m.label}</li>`).join('');
      html += `
        <div class="prix-auto-warning">
          <div class="paw-title">⚠️ Tarifs manquants — non comptés dans le prix</div>
          <ul>${list}</ul>
          <div class="paw-hint">Ouvre <strong>⚙️ Tarifs &amp; Marges</strong> pour ajouter ces items au tarif et obtenir un prix complet.</div>
        </div>
      `;
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

      base.jours = card.querySelector('[name="nbJours"]')?.value || '1';
      base.tache = card.querySelector('[name="tache"]')?.value || '';
      base.rdvs = Array.from(card.querySelectorAll('.rdv-row')).map(r => ({
        date: r.querySelector('[name="posteDate"]')?.value || '',
        heure: r.querySelector('[name="heureRDV"]')?.value || '8H00'
      }));

      // Engins : dropdown
      const enginsRows = card.querySelector('.detail-section.engins .rows');
      base.engins = this._collectRows(enginsRows, true);
      // Personnel, Véhicules, Matériel : checkboxes (avec multi-lignes durée)
      ['personnel', 'vehicules', 'materiel'].forEach(type => {
        const entries = [];
        card.querySelectorAll(`.cb-grid[data-type="${type}"] input[type="checkbox"]:checked`).forEach(cb => {
          const cbItem = cb.closest('.cb-item');
          const rows = cbItem.querySelectorAll('.cb-row');
          if (rows.length > 0) {
            // Personnel / Véhicules : plusieurs lignes qté+durée possibles
            rows.forEach(row => {
              const qty = row.querySelector('.cb-qty')?.value || '1';
              const duree = row.querySelector('.cb-duree')?.value || '';
              const entry = `${qty}x ${cb.value}`;
              entries.push(duree ? `${entry} [${duree}]` : entry);
            });
          } else {
            // Matériel : juste qté
            const qty = cbItem.querySelector('.cb-qty')?.value || '1';
            entries.push(`${qty}x ${cb.value}`);
          }
        });
        base[type] = entries;
      });
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
        const duree = r.querySelector('[name="enginDuree"]')?.value || '';
        const base = `${qty}x ${label} (${ton}T)`;
        return duree ? `${base} [${duree}]` : base;
      }
      return `${qty}x ${label}`;
    }).filter(s => s);
  },

  loadPostes(postes) {
    this._container.innerHTML = '';
    (postes || []).forEach(p => {
      this.addPoste('detail');
      const card = this._container.lastElementChild;
      card.querySelector('[name="posteTitre"]').value = p.titre || '';
      card.querySelector('[name="postePrix"]').value = p.prix || '';

      // Nombre de jours
      const nbJoursInput = card.querySelector('[name="nbJours"]');
      if (nbJoursInput) nbJoursInput.value = p.jours || '1';

      // Instructions / tâche
      const tacheInput = card.querySelector('[name="tache"]');
      if (tacheInput) tacheInput.value = p.tache || p.text || '';

      // RDV (dates et heures)
      const rdvContainer = card.querySelector('.rdv-container');
      const rdvs = (p.rdvs || []).filter(r => r.date);
      if (rdvs.length > 0 && rdvContainer) {
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

      // Engins (dropdown avec tonnage)
      const enginsRows = card.querySelector('.detail-section.engins .rows');
      const enginsItems = p.engins || [];
      enginsItems.forEach(itemStr => {
        const match = String(itemStr).match(/^(\d+)x\s+(.+?)(?:\s+\((\d+)T\))?(?:\s+\[(.+?)\])?$/);
        if (!match) return;
        const qty = match[1], label = match[2].trim(), ton = match[3] || '', duree = match[4] || '';
        const row = this._createRow('engins');
        const select = row.querySelector('select');
        const qtyInput = row.querySelector('input[name="qty"]');
        if (select) {
          const option = Array.from(select.options).find(o => o.value === label);
          if (option) { select.value = label; }
          else { CustomItems.add('engins', label); select.innerHTML = this._createOptionsHTML('engins'); select.value = label; }
        }
        if (qtyInput) qtyInput.value = qty;
        if (ton) { const tonInput = row.querySelector('input[name="ton"]'); if (tonInput) tonInput.value = ton; }
        if (duree) { const dureeSelect = row.querySelector('[name="enginDuree"]'); if (dureeSelect) dureeSelect.value = duree; }
        enginsRows.appendChild(row);
      });

      // Personnel, Véhicules, Matériel (checkboxes avec multi-lignes)
      ['personnel', 'vehicules', 'materiel'].forEach(type => {
        const items = p[type] || [];
        if (items.length === 0) return;
        const grid = card.querySelector(`.cb-grid[data-type="${type}"]`);
        if (!grid) return;

        // Grouper par label pour permettre plusieurs durées sur la même ressource
        const grouped = {};
        items.forEach(itemStr => {
          const match = String(itemStr).match(/^(\d+)x\s+(.+?)(?:\s+\[(.+?)\])?$/);
          if (!match) return;
          const qty = match[1], label = match[2].trim(), duree = match[3] || '';
          if (!grouped[label]) grouped[label] = [];
          grouped[label].push({ qty: parseInt(qty) || 1, duree });
        });

        Object.entries(grouped).forEach(([label, entries]) => {
          const existing = grid.querySelector(`input[type="checkbox"][value="${CSS.escape(label)}"]`);
          if (existing) {
            existing.checked = true;
            const cbItem = existing.closest('.cb-item');
            cbItem.classList.add('active');
            const rowsContainer = cbItem.querySelector('.cb-rows');
            if (rowsContainer) {
              // Personnel/Véhicules : restaurer les lignes durée
              rowsContainer.classList.remove('hidden');
              rowsContainer.innerHTML = '';
              entries.forEach((e, i) => PosteManager._addDureeRow(rowsContainer, e.qty, e.duree, i === 0));
            } else {
              // Matériel : juste la qté
              const qtyInput = cbItem.querySelector('.cb-qty');
              if (qtyInput) { qtyInput.value = entries[0].qty; qtyInput.classList.remove('hidden'); }
            }
          } else {
            // Item personnalisé pas dans les défauts
            CustomItems.add(type, label);
            const el = this._addCheckboxItem(grid, label, true, entries[0].qty, type, entries[0].duree);
            // Ajouter les lignes supplémentaires
            if (entries.length > 1) {
              const rowsContainer = el.querySelector('.cb-rows');
              if (rowsContainer) {
                entries.slice(1).forEach(e => PosteManager._addDureeRow(rowsContainer, e.qty, e.duree, false));
              }
            }
          }
        });
      });
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

    // Le bouton "listBtn" dans la sidebar refresh la liste
    document.getElementById('listBtn')?.addEventListener('click', () => this.refresh());
    document.getElementById('closeDossierList')?.addEventListener('click', () => this.close());
    this._filter?.addEventListener('input', () => this._render());

    // Charge automatiquement au démarrage si utilisateur connecté
    if (UserManager.getUserId && UserManager.getUserId()) {
      this.fetch();
    }
  },

  toggle() { this.refresh(); },

  async open() {
    if (!this._loaded) await this.fetch();
    else this._render();
  },

  close() {
    // Plus de panel à cacher dans la nouvelle UI
  },

  async fetch() {
    this._body.innerHTML = '<div class="dossier-loading">Chargement des dossiers...</div>';
    try {
      const url = `${CONFIG.SCRIPT_URL}?action=lister&user=${encodeURIComponent(UserManager.getUserId())}`;
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
          (d.client || '').toLowerCase().includes(query)
        )
      : this._dossiers;

    if (filtered.length === 0) {
      this._body.innerHTML = `<div class="side-empty">${query ? 'Aucun résultat' : 'Aucun dossier'}</div>`;
      return;
    }

    const fmtCHF = (v) => {
      const n = parseFloat(v) || 0;
      if (n >= 1000) return (n / 1000).toFixed(n >= 10000 ? 0 : 1) + 'k';
      return Math.round(n).toString();
    };

    const activeRef = document.querySelector('[name="ref"]')?.value?.trim();

    const statusFromString = (s) => {
      const lo = (s || '').toLowerCase();
      if (lo.includes('refus') || lo.includes('non accept') || lo.includes('annul')) return 'status-refus';
      if (lo.includes('brouillon')) return 'status-draft';
      if (lo.includes('réalis') || lo.includes('realis')) return 'status-won';
      if (lo.includes('ok') || lo.includes('accept')) return 'status-ok';
      return 'status-draft';
    };

    this._body.innerHTML = filtered.map(d => {
      const statusClass = statusFromString(d.statut);
      const isActive = d.ref === activeRef;
      return `
        <div class="dossier-item ${isActive ? 'active' : ''}" data-ref="${d.ref}" data-statut="${d.statut || ''}">
          <span class="dossier-dot ${statusClass}" title="${d.statut || ''}"></span>
          <div class="dossier-ref">${d.ref}</div>
          <div class="dossier-amount">${fmtCHF(d.montantHT)}</div>
          <div class="dossier-client">${d.client || '—'}</div>
          <div class="dossier-meta">${(d.date || '').split(' ')[0]}</div>
          <button type="button" class="dossier-menu-btn" data-ref="${d.ref}" aria-label="Actions">
            <svg width="12" height="12" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="8" cy="3" r="1.2"/><circle cx="8" cy="8" r="1.2"/><circle cx="8" cy="13" r="1.2"/></svg>
          </button>
        </div>
      `;
    }).join('');

    this._body.querySelectorAll('.dossier-item').forEach(row => {
      row.addEventListener('click', (e) => {
        // Si clic sur le menu, ne pas charger
        if (e.target.closest('.dossier-menu-btn')) return;
        const ref = row.dataset.ref;
        const sr = document.getElementById('searchRef');
        if (sr) sr.value = ref;
        DossierLoader.load(ref);
      });
    });

    // Menu contextuel — clic sur le bouton ⋮
    this._body.querySelectorAll('.dossier-menu-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        e.preventDefault();
        DossierMenu.show(btn);
      });
    });
    // Right-click sur la ligne du dossier ouvre aussi le menu
    this._body.querySelectorAll('.dossier-item').forEach(row => {
      row.addEventListener('contextmenu', (e) => {
        e.preventDefault();
        const btn = row.querySelector('.dossier-menu-btn');
        if (btn) DossierMenu.showAt(btn, e.clientX, e.clientY);
      });
    });
  }
};

// ============================================
// DOSSIER CONTEXT MENU (statut + delete)
// ============================================
const DossierMenu = {
  _menu: null,
  _currentRef: null,

  init() {
    document.addEventListener('click', () => this.hide());
  },

  show(btn) {
    const r = btn.getBoundingClientRect();
    this.showAt(btn, r.left, r.bottom + 4);
  },

  showAt(btn, x, y) {
    this._currentRef = btn.dataset.ref;
    this.hide();
    const menu = document.createElement('div');
    menu.className = 'dossier-context-menu';
    menu.innerHTML = `
      <button type="button" data-action="status:Accepté"><span class="dot status-ok"></span> Marquer accepté</button>
      <button type="button" data-action="status:Réalisé"><span class="dot status-won"></span> Marquer réalisé</button>
      <button type="button" data-action="status:Refusé"><span class="dot status-refus"></span> Marquer refusé</button>
      <button type="button" data-action="status:Brouillon"><span class="dot status-draft"></span> Repasser brouillon</button>
      <hr>
      <button type="button" data-action="delete" class="danger">🗑️ Supprimer</button>
    `;
    document.body.appendChild(menu);
    menu.style.top = Math.min(y, window.innerHeight - menu.offsetHeight - 16) + 'px';
    menu.style.left = Math.min(x, window.innerWidth - 220) + 'px';
    menu.querySelectorAll('button').forEach(b => {
      b.addEventListener('click', (e) => {
        e.stopPropagation();
        const action = b.dataset.action;
        this._do(action);
        this.hide();
      });
    });
    this._menu = menu;
    setTimeout(() => menu.classList.add('open'), 10);
  },

  hide() {
    if (this._menu) { this._menu.remove(); this._menu = null; }
  },

  async _do(action) {
    const ref = this._currentRef;
    if (!ref) return;

    if (action === 'delete') {
      if (!confirm(`Supprimer définitivement le dossier ${ref} ?\n(toutes les versions seront retirées du suivi)`)) return;
      try {
        const resp = await fetch(CONFIG.SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({
            _action: 'delete_dossier',
            ref,
            userId: UserManager.getUserId()
          })
        });
        const res = await resp.json();
        if (res.status === 'success') {
          Toast.success(`🗑️ Dossier ${ref} supprimé (${res.deleted || 1} ligne)`);
          DossierList._loaded = false;
          DossierList.fetch();
        } else {
          Toast.error('Erreur : ' + (res.message || 'Suppression échouée'));
        }
      } catch (err) {
        Toast.error('Erreur : ' + err.message);
      }
      return;
    }

    if (action.startsWith('status:')) {
      const statut = action.split(':')[1];
      try {
        const resp = await fetch(CONFIG.SCRIPT_URL, {
          method: 'POST',
          body: JSON.stringify({
            _action: 'set_status',
            ref,
            statut,
            userId: UserManager.getUserId()
          })
        });
        const res = await resp.json();
        if (res.status === 'success') {
          Toast.success(`✅ ${ref} → ${statut}`);
          DossierList._loaded = false;
          DossierList.fetch();
        } else {
          Toast.error('Erreur : ' + (res.message || 'Statut non modifié'));
        }
      } catch (err) {
        Toast.error('Erreur : ' + err.message);
      }
    }
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
      const url = `${CONFIG.SCRIPT_URL}?action=rechercher&ref=${encodeURIComponent(ref.trim())}&user=${encodeURIComponent(UserManager.getUserId())}`;
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
      // Mémoriser le réalisé existant pour restoration dans le modal
      window._currentDossierRealise = data.realise || null;
      PriceCalc.updateBreakdown();
      Toast.success('Dossier chargé !' + (data.realise ? ' (avec réalisé)' : ''));
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
    data.userId = UserManager.getUserId();

    Modal.show(data, () => this._doSubmit(data));
  },

  /** Sauvegarde sans génération de devis/RESA (brouillon dans le Spreadsheet) */
  async saveDraft(form) {
    if (this._submitting) return;
    // Validation minimale pour brouillon : juste référence requise
    const refInput = form.querySelector('[name="ref"]');
    const ref = (refInput?.value || '').trim();
    if (!ref) {
      Toast.error('Référence requise pour sauvegarder un brouillon');
      refInput?.focus();
      return;
    }

    const fd = new FormData(form);
    const data = Object.fromEntries(fd.entries());
    data.postes = PosteManager.collectAll();
    data.userId = UserManager.getUserId();
    data._action = 'save_draft';

    const btn = document.getElementById('saveDraftBtn');
    if (btn) { btn.disabled = true; btn.textContent = '⏳ Sauvegarde…'; }

    try {
      const resp = await fetch(CONFIG.SCRIPT_URL, {
        method: 'POST',
        body: JSON.stringify(data)
      });
      const res = await resp.json();
      if (res.status === 'success') {
        Toast.success(`💾 Brouillon "${ref}" sauvegardé`);
        // Rafraîchir la sidebar dossiers
        if (typeof DossierList !== 'undefined') {
          DossierList._loaded = false;
          DossierList.fetch();
        }
      } else {
        Toast.error('Erreur : ' + (res.message || 'Sauvegarde échouée'));
      }
    } catch (err) {
      Toast.error('Erreur de connexion : ' + err.message);
    } finally {
      if (btn) {
        btn.disabled = false;
        btn.innerHTML = `<svg width="12" height="12" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.7" style="margin-right:2px"><path d="M3 5a2 2 0 012-2h7l5 5v9a2 2 0 01-2 2H5a2 2 0 01-2-2V5zM12 3v5h5M7 14h6M7 11h6" stroke-linecap="round" stroke-linejoin="round"/></svg> Brouillon`;
      }
    }
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

        // Téléchargements automatiques : PDF devis + DOCX devis + PDF RESA + DOCX RESA
        const downloads = [
          { b64: res.pdfB64, name: res.pdfName, mime: 'application/pdf' },
          { b64: res.docxB64, name: res.docxName, mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
          { b64: res.resaPdfB64, name: res.resaPdfName, mime: 'application/pdf' },
          { b64: res.resaDocxB64, name: res.resaDocxName, mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
        ].filter(d => d.b64 && d.name);

        // Déclencher chaque téléchargement espacé de 600 ms (certains navigateurs bloquent les téléchargements simultanés)
        downloads.forEach((d, i) => {
          setTimeout(() => this._downloadBase64File(d.b64, d.name, d.mime), i * 600);
        });

        // WhatsApp désactivé temporairement
        // const waNumber = (data.mobileSociete || '').replace(/\s+/g, '');
        // if (waNumber) {
        //   const waText = encodeURIComponent(`Documents ${data.ref}:\nPDF: ${res.pdfUrl}`);
        //   setTimeout(() => window.open(`https://wa.me/${waNumber}?text=${waText}`, '_blank'), downloads.length * 600 + 500);
        // }
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
  },

  /** Décode un fichier base64 et déclenche un téléchargement automatique */
  _downloadBase64File(b64, filename, mime) {
    try {
      const binary = atob(b64);
      const bytes = new Uint8Array(binary.length);
      for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
      const blob = new Blob([bytes], { type: mime || 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1500);
      Toast.info('📥 ' + filename);
    } catch (err) {
      Toast.error('Erreur téléchargement ' + filename + ' : ' + err.message);
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
// VENTILATION EXPORT (génération Excel)
// ============================================
const VentilationExport = {

  async generate() {
    // Collecter les données du formulaire
    const form = document.getElementById('devisForm');
    const fd = new FormData(form);
    const data = Object.fromEntries(fd.entries());
    data.postes = PosteManager.collectAll();
    data.userId = UserManager.getUserId();

    const postsDetail = data.postes.filter(p => p.mode === 'detail');
    if (postsDetail.length === 0) {
      Toast.warning('Aucun poste logistique détaillé à ventiler.');
      return;
    }

    Toast.info('Génération de la ventilation...');

    try {
      const url = `${CONFIG.SCRIPT_URL}?action=ventilation&data=${encodeURIComponent(JSON.stringify(data))}`;
      const resp = await fetch(url);
      const result = await resp.json();

      if (result.status !== 'success' || !result.data) {
        Toast.error('Erreur ventilation : ' + (result.message || 'Inconnue'));
        return;
      }

      // Ajouter les lignes de supplément km si calculé
      let lignes = result.data;
      const excessKm = KmCalculator.getExcessKm();
      if (excessKm > 0) {
        // Trouver les véhicules utilisés dans les postes pour calculer le coût km
        const vehiculesUtilises = new Set();
        data.postes.forEach(p => {
          (p.vehicules || []).forEach(v => {
            const m = String(v).match(/^\d+x\s+(.+?)(?:\s+\[.+?\])?$/);
            if (m) vehiculesUtilises.add(m[1].trim());
          });
        });

        // Insérer les lignes km AVANT le total (dernière ligne)
        const totalLine = lignes.pop();
        let n = totalLine.n;
        vehiculesUtilises.forEach(veh => {
          const rate = KmCalculator.KM_RATES[veh];
          if (!rate) return;
          const cost = excessKm * rate;
          lignes.push({
            n: n++,
            ventil: 'LOC5',
            poste: 'FR-4300-01',
            libelle: `Suppl. km ${veh} : ${excessKm} km x ${rate} CHF/km`,
            montant: cost,
            dev: 'CHF'
          });
        });

        // Recalculer le total
        totalLine.n = n++;
        totalLine.montant = lignes.reduce((s, l) => s + (l.montant || 0), 0);
        lignes.push(totalLine);
      }

      this._downloadExcel(lignes, data.ref || 'DEVIS');
      Toast.success('Fichier Excel ventilation téléchargé !');
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    }
  },

  _downloadExcel(lignes, ref) {
    // Construire le CSV (compatible Excel avec séparateur ; pour la locale FR)
    const BOM = '\uFEFF';
    const headers = ['N°', 'Ventil', 'Poste', 'Libellé', 'Sst.', 'Montant', 'Dev', 'Mnt Réel', 'Dev', 'Mnt Reçu', 'Dev', 'TVA'];
    let csv = BOM + headers.join(';') + '\n';

    lignes.forEach(l => {
      const montantFmt = (l.montant || 0).toFixed(2).replace('.', ',');
      const row = [
        l.n || '',
        l.ventil || '',
        l.poste || '',
        '"' + (l.libelle || '').replace(/"/g, '""') + '"',
        '', // Sst
        montantFmt,
        l.dev || 'CHF',
        '', // Mnt Réel
        'CHF',
        '', // Mnt Reçu
        'CHF',
        '0'
      ];
      csv += row.join(';') + '\n';
    });

    // Télécharger
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Ventilation_${ref}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }
};

// ============================================
// PLANNING CONSOLIDÉ
// ============================================
const PlanningExport = {

  async generate() {
    const userId = UserManager.getUserId();
    if (!userId) {
      Toast.error('Veuillez vous connecter d\'abord.');
      return;
    }

    const btn = document.getElementById('planningBtn');
    if (btn) { btn.disabled = true; btn.textContent = 'Génération en cours...'; }

    Toast.info('Génération du planning consolidé...');

    try {
      const url = `${CONFIG.SCRIPT_URL}?action=planning_generate&user=${encodeURIComponent(userId)}`;
      const resp = await fetch(url);
      const result = await resp.json();

      if (result.status === 'success' && result.url) {
        Toast.success(`Planning généré ! ${result.nbClients} client(s), ${result.nbJours} jours.`);
        window.open(result.url, '_blank');
      } else {
        Toast.error('Erreur : ' + (result.message || 'Inconnue'));
      }
    } catch (err) {
      Toast.error('Erreur : ' + err.message);
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = '📅 Générer Planning Consolidé'; }
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
  // 1. Charger les utilisateurs + donnees en parallele
  await Promise.all([
    UserManager.fetchUsers(),
    CustomItems.fetchFromServer(),
    TarifManager.fetchFromServer()
  ]);

  // 2. Afficher l'ecran de login (ou auto-login si deja sauvegarde)
  await UserManager.showLogin();

  // 3. Init des modules
  SyncManager.init();
  PosteManager.init();
  PriceCalc.init();
  DossierList.init();
  SettingsPanel.init();
  ProfilePanel.init();
  CompanySearch.init();
  KmCalculator.init();
  RightRail.init();
  MobileShell.init();
  PWAInstall.init();
  RealiseManager.init();
  DossierMenu.init();
  AffairesView.init();
  ApiKeysManager.init();
  CalendarView.init();

  // Logout
  document.getElementById('logoutBtn')?.addEventListener('click', () => {
    DossierList._loaded = false;
    UserManager.logout();
  });

  // Search bar
  document.getElementById('loadBtn')?.addEventListener('click', () => {
    DossierLoader.load(document.getElementById('searchRef').value);
  });
  document.getElementById('searchRef')?.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); DossierLoader.load(e.target.value); }
  });

  // Auto total button
  document.getElementById('autoTotalBtn')?.addEventListener('click', () => PriceCalc.applyAutoTotal());
  document.getElementById('saveDraftBtn')?.addEventListener('click', () => {
    const form = document.getElementById('devisForm');
    if (form) FormSubmitter.saveDraft(form);
  });

  // Export ventilation
  document.getElementById('exportVentilBtn')?.addEventListener('click', () => VentilationExport.generate());

  // Planning consolidé
  document.getElementById('planningBtn')?.addEventListener('click', () => PlanningExport.generate());

  // Form submit
  document.getElementById('devisForm')?.addEventListener('submit', (e) => {
    e.preventDefault();
    FormSubmitter.submit(e.target);
  });
});
