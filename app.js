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
    const customs = (CustomItems.load()[type] || []).filter(c => !defaults.includes(c));
    const allItems = [...defaults, ...customs];
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
    data.userId = UserManager.getUserId();

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

      this._downloadExcel(result.data, data.ref || 'DEVIS');
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
