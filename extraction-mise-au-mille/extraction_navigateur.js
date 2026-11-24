/*
 * Extraction "Mise au mille" - version navigateur, sans installation.
 *
 * Ce script s'execute DANS la page de l'appli (favori ou console F12 de Edge) :
 * il regle la date jour apres jour, declenche l'extraction Excel, garde chaque
 * fichier en memoire du navigateur, puis propose un seul ZIP contenant tous les
 * fichiers renommes avec la date du programme.
 * Mode d'emploi : ouvrir extraction_navigateur.html.
 */
(() => {
  'use strict';
  if (window.__mam) { window.__mam.montrer(); return; }

  const VERSION = '1';
  const REGLAGES = {
    debut: '2025-07-01',        // premier jour propose
    prefixe: 'Mise_au_mille_',  // nom des fichiers : prefixe + AAAA-MM-JJ + extension
    exporterSiVide: false,      // true : extraire aussi les jours ou le tableau est vide
    attenteExport: 120,         // secondes max pour qu'un fichier arrive apres le clic
    silenceSansExport: 15,      // secondes sans aucune activite apres le clic -> "aucun fichier"
    attenteChargement: 90,      // secondes max de chargement de la page pour un jour
    erreursDeSuiteMax: 5,
    identiquesMax: 4,
  };
  const FOND_IGNORE = /signalr|\/hubs?\/|socket\.io|sockjs|negotiate|\/poll/i;
  const BOUTONS = 'button, a, [role=button], input[type=button], input[type=submit]';
  const GRILLES = '.dx-datagrid, .dx-treelist, .dx-pivotgrid';

  // ------------------------------------------------------------------ outils
  const deux = n => String(n).padStart(2, '0');
  const isoLocal = d => d.getFullYear() + '-' + deux(d.getMonth() + 1) + '-' + deux(d.getDate());
  const versDate = iso => { const [a, m, j] = iso.split('-').map(Number); return new Date(a, m - 1, j); };
  const jourFr = iso => iso.split('-').reverse().join('/');
  const horsPanneau = el => !el.closest('[data-bot-barre]');
  const visible = el => {
    const r = el.getBoundingClientRect(), s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  const texteCourt = el => (el.innerText || el.value || '').trim().replace(/\s+/g, ' ');
  const q = v => '"' + String(v).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';

  // Fonctions d'origine du navigateur, gardees avant de poser les intercepteurs.
  const natif = {
    fetch: window.fetch,
    clicLien: HTMLAnchorElement.prototype.click,
    envoyer: EventTarget.prototype.dispatchEvent,
    creerUrl: URL.createObjectURL,
    valeur: Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set,
  };

  // Pause : passe par un Worker quand c'est possible, car les minuteries d'un onglet
  // masque sont fortement ralenties par le navigateur.
  let dormir = ms => new Promise(r => setTimeout(r, ms));
  try {
    const source = natif.creerUrl.call(URL, new Blob(
      ['onmessage = e => setTimeout(() => postMessage(e.data.n), e.data.ms);'], {type: 'text/javascript'}));
    const ouvrier = new Worker(source);
    const attentes = new Map();
    let n = 0;
    const repli = () => { dormir = ms => new Promise(r => setTimeout(r, ms)); attentes.forEach(r => r()); attentes.clear(); };
    ouvrier.onmessage = e => { const r = attentes.get(e.data); attentes.delete(e.data); if (r) r(); };
    ouvrier.onerror = repli;
    dormir = ms => new Promise(r => { attentes.set(++n, r); ouvrier.postMessage({n, ms}); });
  } catch (e) { /* minuterie classique */ }

  // Clic complet (appui, relachement, clic), comme le ferait la souris.
  function clicSimule(el, sansClic) {
    const r = el.getBoundingClientRect();
    const o = {bubbles: true, cancelable: true, view: window, button: 0,
               clientX: r.left + r.width / 2, clientY: r.top + r.height / 2};
    const p = {...o, pointerId: 1, isPrimary: true, pointerType: 'mouse'};
    natif.envoyer.call(el, new PointerEvent('pointerdown', {...p, buttons: 1}));
    natif.envoyer.call(el, new MouseEvent('mousedown', {...o, buttons: 1}));
    natif.envoyer.call(el, new PointerEvent('pointerup', p));
    natif.envoyer.call(el, new MouseEvent('mouseup', o));
    if (!sansClic) natif.envoyer.call(el, new MouseEvent('click', o));
  }

  // ---------------------------------------------------- activite du reseau
  const reseau = {
    enCours: new Map(), dernier: performance.now(), jeton: 0,
    debut(url) {
      if (FOND_IGNORE.test(String(url || ''))) return 0;
      this.enCours.set(++this.jeton, performance.now());
      this.dernier = performance.now();
      return this.jeton;
    },
    fin(jeton) { if (this.enCours.delete(jeton)) this.dernier = performance.now(); },
    calmeDepuis() {
      const maintenant = performance.now();
      // une requete en attente depuis plus de 2 minutes est une connexion de fond
      this.enCours.forEach((depart, jeton) => { if (maintenant - depart > 120000) this.enCours.delete(jeton); });
      return this.enCours.size ? 0 : maintenant - this.dernier;
    },
    marquer() { this.dernier = performance.now(); },
  };
  const xhrOuvrir = XMLHttpRequest.prototype.open, xhrEnvoyer = XMLHttpRequest.prototype.send;
  XMLHttpRequest.prototype.open = function (methode, url) { this.__mamUrl = url; return xhrOuvrir.apply(this, arguments); };
  XMLHttpRequest.prototype.send = function () {
    const jeton = reseau.debut(this.__mamUrl);
    if (jeton) this.addEventListener('loadend', () => reseau.fin(jeton), {once: true});
    return xhrEnvoyer.apply(this, arguments);
  };
  if (natif.fetch) {
    window.fetch = function (ressource) {
      const jeton = reseau.debut(ressource && ressource.url ? ressource.url : ressource);
      const promesse = natif.fetch.apply(this, arguments);
      if (jeton) promesse.then(() => reseau.fin(jeton), () => reseau.fin(jeton));
      return promesse;
    };
  }
  const chargementVisible = () => Array.from(document.querySelectorAll('.dx-loadpanel-content')).some(visible);

  async function calme(silence) {
    const debut = performance.now();
    await dormir(250);
    while (performance.now() - debut < REGLAGES.attenteChargement * 1000) {
      if (reseau.calmeDepuis() >= (silence || 700) && !chargementVisible()) return;
      await dormir(120);
    }
    throw new Error('la page charge toujours');
  }

  // ------------------------------------- capture du fichier au lieu du telechargement
  let capture = null;
  function intercepter(lien) {
    if (!capture || !(lien instanceof HTMLAnchorElement)) return false;
    const adresse = lien.href || '';
    if (!lien.hasAttribute('download') && !/^(blob|data):/.test(adresse)) return false;
    const c = capture;
    capture = null;
    (async () => {
      let blob = null;
      try {
        const reponse = await natif.fetch.call(window, adresse, {credentials: 'include'});
        if (reponse.ok) blob = await reponse.blob();
      } catch (e) { /* on se rabat sur le dernier Blob cree */ }
      c.resoudre({nom: lien.getAttribute('download') || '', blob: blob || c.dernierBlob});
    })();
    return true;
  }
  HTMLAnchorElement.prototype.click = function () { if (!intercepter(this)) return natif.clicLien.apply(this, arguments); };
  EventTarget.prototype.dispatchEvent = function (evenement) {
    if (capture && evenement && evenement.type === 'click' && intercepter(this)) return true;
    return natif.envoyer.apply(this, arguments);
  };
  URL.createObjectURL = function (objet) {
    if (capture && objet instanceof Blob) capture.dernierBlob = objet;
    return natif.creerUrl.apply(this, arguments);
  };

  // Clique sur le bouton d'extraction et attend le fichier. null s'il n'en vient aucun.
  async function exporter(bouton) {
    let recu = null;
    capture = {resoudre: f => { recu = f; }, dernierBlob: null};
    reseau.marquer();
    clicSimule(bouton);
    const debut = performance.now();
    while (!recu) {
      await dormir(150);
      const ecoule = (performance.now() - debut) / 1000;
      if (ecoule > REGLAGES.attenteExport) { capture = null; throw new Error('pas de fichier après ' + Math.round(ecoule) + ' s'); }
      if (capture && ecoule > REGLAGES.silenceSansExport && reseau.calmeDepuis() / 1000 > REGLAGES.silenceSansExport
          && !chargementVisible()) { capture = null; return null; }
    }
    if (!recu.blob || !recu.blob.size) throw new Error('fichier vide');
    return recu;
  }

  // ------------------------------------------------ stockage (IndexedDB, sinon memoire)
  const base = {
    db: null, memoire: new Map(), pret: null,
    ouvrir() {
      if (!this.pret) this.pret = new Promise(fini => {
        try {
          const r = indexedDB.open('mam-extraction', 1);
          r.onupgradeneeded = () => r.result.createObjectStore('jours', {keyPath: 'date'});
          r.onsuccess = () => { this.db = r.result; fini(); };
          r.onerror = r.onblocked = () => fini();
        } catch (e) { fini(); }
      });
      return this.pret;
    },
    operation(mode, action) {
      return new Promise((ok, echec) => {
        const t = this.db.transaction('jours', mode);
        const requete = action(t.objectStore('jours'));
        t.oncomplete = () => ok(requete && requete.result);
        t.onerror = t.onabort = () => echec(t.error || new Error('stockage refusé'));
      });
    },
    async ecrire(fiche) {
      await this.ouvrir();
      if (this.db) {
        try { await this.operation('readwrite', s => s.put(fiche)); return; } catch (e) { this.db = null; }
      }
      this.memoire.set(fiche.date, fiche);
    },
    async toutes() {
      await this.ouvrir();
      const fiches = this.db ? await this.operation('readonly', s => s.getAll()) : [];
      return fiches.concat(Array.from(this.memoire.values())).sort((a, b) => a.date < b.date ? -1 : 1);
    },
    async vider() {
      await this.ouvrir();
      this.memoire.clear();
      if (this.db) await this.operation('readwrite', s => s.clear());
    },
  };

  // --------------------------------------------------------------------- ZIP
  const TABLE_CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(octets) {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < octets.length; i++) c = TABLE_CRC[(c ^ octets[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  // Archive ZIP sans compression (les .xlsx sont deja compresses).
  async function construireZip(fichiers) {
    const encodeur = new TextEncoder(), morceaux = [], repertoire = [];
    const d = new Date();
    const heure = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
    const jour = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    let position = 0, tailleRepertoire = 0;
    for (const f of fichiers) {
      const nom = encodeur.encode(f.nom);
      const donnees = new Uint8Array(await f.blob.arrayBuffer());
      const crc = crc32(donnees);
      const local = new DataView(new ArrayBuffer(30));
      local.setUint32(0, 0x04034b50, true); local.setUint16(4, 20, true); local.setUint16(6, 0x0800, true);
      local.setUint16(8, 0, true); local.setUint16(10, heure, true); local.setUint16(12, jour, true);
      local.setUint32(14, crc, true); local.setUint32(18, donnees.length, true); local.setUint32(22, donnees.length, true);
      local.setUint16(26, nom.length, true); local.setUint16(28, 0, true);
      morceaux.push(local.buffer, nom, donnees);
      const central = new DataView(new ArrayBuffer(46));
      central.setUint32(0, 0x02014b50, true); central.setUint16(4, 20, true); central.setUint16(6, 20, true);
      central.setUint16(8, 0x0800, true); central.setUint16(10, 0, true); central.setUint16(12, heure, true);
      central.setUint16(14, jour, true); central.setUint32(16, crc, true); central.setUint32(20, donnees.length, true);
      central.setUint32(24, donnees.length, true); central.setUint16(28, nom.length, true);
      central.setUint32(42, position, true);
      repertoire.push(central.buffer, nom);
      position += 30 + nom.length + donnees.length;
      tailleRepertoire += 46 + nom.length;
    }
    const fin = new DataView(new ArrayBuffer(22));
    fin.setUint32(0, 0x06054b50, true); fin.setUint16(8, fichiers.length, true); fin.setUint16(10, fichiers.length, true);
    fin.setUint32(12, tailleRepertoire, true); fin.setUint32(16, position, true);
    return new Blob(morceaux.concat(repertoire, [fin.buffer]), {type: 'application/zip'});
  }

  // Empreinte du contenu d'un .xlsx (un ZIP) : les CRC de ses parties, hors metadonnees.
  async function empreinte(blob) {
    const o = new Uint8Array(await blob.arrayBuffer()), v = new DataView(o.buffer);
    let fin = -1;
    for (let i = o.length - 22; i >= Math.max(0, o.length - 65557); i--) if (v.getUint32(i, true) === 0x06054b50) { fin = i; break; }
    if (fin < 0) return 'brut:' + o.length + ':' + crc32(o);
    const decodeur = new TextDecoder(), parties = [];
    let p = v.getUint32(fin + 16, true);
    for (let k = v.getUint16(fin + 10, true); k > 0 && p + 46 <= o.length && v.getUint32(p, true) === 0x02014b50; k--) {
      const lNom = v.getUint16(p + 28, true);
      const nom = decodeur.decode(o.subarray(p + 46, p + 46 + lNom));
      if (!nom.startsWith('docProps/')) parties.push(nom + ':' + v.getUint32(p + 16, true));
      p += 46 + lNom + v.getUint16(p + 30, true) + v.getUint16(p + 32, true);
    }
    return parties.sort().join('|');
  }

  // --------------------------------------------- reperage des elements de la page
  const idStable = id => !!id && !/^(mat|cdk|pr_id|ng|el-id|input-|v-|radix|headlessui|mui|:r)/i.test(id) &&
    !/\d{3,}/.test(id) && !/[-_:]\d+$/.test(id);

  function chemin(el) {
    const morceaux = [];
    for (let e = el; e && e.nodeType === 1 && e !== document.body && e !== document.documentElement; e = e.parentElement) {
      if (idStable(e.id)) { morceaux.unshift('#' + CSS.escape(e.id)); return morceaux.join(' > '); }
      let n = 1;
      for (let s = e.previousElementSibling; s; s = s.previousElementSibling) if (s.tagName === e.tagName) n++;
      morceaux.unshift(e.tagName.toLowerCase() + ':nth-of-type(' + n + ')');
    }
    morceaux.unshift('body');
    return morceaux.join(' > ');
  }

  // Maniere de retrouver un element plus tard (la page peut le reconstruire entre deux jours).
  function repere(el) {
    const tag = el.tagName.toLowerCase(), essais = [];
    if (idStable(el.id)) essais.push('#' + CSS.escape(el.id));
    for (const a of ['data-testid', 'data-test', 'data-cy', 'formcontrolname', 'name', 'aria-label', 'placeholder', 'title']) {
      const v = el.getAttribute(a);
      if (v) essais.push(tag + '[' + a + '=' + q(v) + ']');
    }
    const boiteDx = tag === 'input' && el.closest('.dx-datebox');
    if (boiteDx) {
      if (idStable(boiteDx.id)) essais.push('#' + CSS.escape(boiteDx.id) + ' input.dx-texteditor-input');
      essais.push('.dx-datebox input.dx-texteditor-input');
    }
    for (const c of el.classList) if (/export/i.test(c)) essais.push('.' + CSS.escape(c));
    essais.push(chemin(el));
    for (const seulementUnique of [true, false]) {
      for (const css of essais) {
        let tous;
        try { tous = Array.from(document.querySelectorAll(css)).filter(horsPanneau); } catch (e) { continue; }
        const rang = tous.indexOf(el);
        if (rang >= 0 && (tous.length === 1 || !seulementUnique)) return {css, rang};
      }
    }
    return null;
  }
  function retrouver(r) {
    if (!r) return null;
    try { return Array.from(document.querySelectorAll(r.css)).filter(horsPanneau)[r.rang] || null; } catch (e) { return null; }
  }

  function scoreDate(el) {
    if (!visible(el) || el.disabled || !horsPanneau(el)) return 0;
    const type = (el.getAttribute('type') || 'text').toLowerCase();
    if (type === 'date') return 100;
    if (!/^(text|search|tel)$/.test(type)) return 0;
    if (el.closest('.dx-datebox')) return el.closest(GRILLES + ', .dx-overlay-content') ? 70 : 95;
    let indices = Array.from(el.attributes).map(a => a.name + '=' + a.value).join(' ');
    for (let e = el.parentElement, i = 0; e && i < 3; e = e.parentElement, i++) indices += ' ' + e.tagName + ' ' + (e.getAttribute('class') || '');
    if (/date-?picker|calendar|flatpickr|date-?input|date-?box|el-date|ant-picker|dp__|p-calendar/i.test(indices)) return 90;
    if (/^\s*(\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}|\d{4}-\d{2}-\d{2})\s*$/.test(el.value || '')) return 80;
    return 0;
  }
  function scoreExport(el) {
    if (!visible(el) || !horsPanneau(el)) return 0;
    const bits = [el.innerText, el.value, el.title, el.getAttribute('aria-label'), el.getAttribute('class'), el.getAttribute('download')];
    el.querySelectorAll('i, svg, img, span, mat-icon').forEach(c => bits.push(c.getAttribute('class'), c.getAttribute('title'), c.getAttribute('alt')));
    const t = bits.filter(Boolean).join(' ').toLowerCase();
    let s = 0;
    if (/extraction|extraire|export/.test(t)) s += 50;
    if (/excel|xlsx?\b/.test(t)) s += 40;
    if (/t[ée]l[ée]charg|download|get_app/.test(t)) s += 30;
    if (s && /pdf|csv|import/.test(t)) s -= 30;
    return s;
  }
  function meilleur(selecteur, score) {
    const xs = Array.from(document.querySelectorAll(selecteur)).map(el => [score(el), el]).filter(x => x[0] > 0).sort((a, b) => b[0] - a[0]);
    return !xs.length || (xs.length > 1 && xs[1][0] === xs[0][0]) ? null : xs[0][1];
  }
  function versInput(el) {
    if (el.tagName === 'INPUT') return el;
    if (el.tagName === 'LABEL' && el.control && el.control.tagName === 'INPUT') return el.control;
    const boiteDx = el.closest('.dx-datebox');
    if (boiteDx) return boiteDx.querySelector('input.dx-texteditor-input');
    for (let e = el, i = 0; e && i < 6; e = e.parentElement, i++) {
      const ins = Array.from(e.querySelectorAll('input')).filter(visible);
      if (ins.length === 1) return ins[0];
      if (ins.length > 1) return null;
    }
    return null;
  }

  // ------------------------------------------------- selecteur de date DevExtreme
  // Vraie valeur d'un dxDateBox (champ cache), independante du format affiche.
  function valeurDx(champ) {
    const racine = champ.closest('.dx-datebox');
    const cache = racine && racine.querySelector('input[type=hidden]');
    if (!cache) return undefined;      // pas un dxDateBox (ou trop ancien)
    if (!cache.value) return null;
    if (/^\d{4}-\d{2}-\d{2}$/.test(cache.value)) return cache.value;
    const d = new Date(cache.value);
    return isNaN(d) ? null : isoLocal(d);
  }

  // Texte a saisir pour un jour, d'apres le texte affiche et la vraie date qu'il represente.
  function texteCommeAffiche(affiche, vraieIso, jourIso) {
    const m = /^(\d{1,4})([\/.\-])(\d{1,2})\2(\d{1,4})$/.exec((affiche || '').trim());
    if (!m || !vraieIso) return null;
    const [a, mo, j] = vraieIso.split('-').map(Number), parts = [m[1], m[3], m[4]].map(Number);
    const ordres = [['j', 'm', 'A'], ['m', 'j', 'A'], ['A', 'm', 'j'], ['j', 'm', 'a'], ['m', 'j', 'a']];
    const vraies = {j, m: mo, A: a, a: a % 100};
    const ordre = ordres.find(o => o.every((cle, i) => parts[i] === vraies[cle]) && (o.includes('A') ? true : m[4].length <= 2));
    if (!ordre) return null;
    const sansZero = [m[1], m[3], m[4]].some((p, i) => ordre[i] !== 'A' && ordre[i] !== 'a' && p.length === 1);
    const [ca, cm, cj] = jourIso.split('-').map(Number);
    const cible = {j: sansZero ? String(cj) : deux(cj), m: sansZero ? String(cm) : deux(cm), A: String(ca), a: deux(ca % 100)};
    return ordre.map(cle => cible[cle]).join(m[2]);
  }

  function saisieSimulee(champ, texte) {
    champ.focus();
    natif.valeur.call(champ, texte);
    natif.envoyer.call(champ, new Event('input', {bubbles: true}));
    natif.envoyer.call(champ, new Event('change', {bubbles: true}));
    champ.blur();
  }

  // Calendrier ouvert : mois affiche et case du jour voulu ("AAAA/MM/JJ").
  function viserJour(vise) {
    const cals = Array.from(document.querySelectorAll('.dx-overlay-wrapper .dx-calendar')).filter(visible);
    if (!cals.length) return {ouvert: false};
    const cal = cals[cals.length - 1];
    const corps = cal.querySelector('.dx-calendar-body').getBoundingClientRect();
    const aLecran = td => { const r = td.getBoundingClientRect(), x = r.left + r.width / 2; return r.width > 0 && x > corps.left && x < corps.right; };
    const cases = Array.from(cal.querySelectorAll('td.dx-calendar-cell:not(.dx-calendar-other-view)')).filter(aLecran);
    const cible = cases.find(td => td.getAttribute('data-value') === vise) || null;
    return {ouvert: true, cal, vueMois: cal.classList.contains('dx-calendar-view-month'),
            mois: cases.length ? (cases[0].getAttribute('data-value') || '').slice(0, 7) : null,
            interdit: !!cible && cible.classList.contains('dx-calendar-empty-cell'), cible};
  }

  async function calendrierDx(champ, jourIso) {
    const vise = jourIso.replace(/-/g, '/'), mois = vise.slice(0, 7);
    if (!viserJour(vise).ouvert) {
      clicSimule(champ.closest('.dx-datebox').querySelector('.dx-dropdowneditor-button') || champ);
      await dormir(500);
    }
    for (let i = 0; i < 80; i++) {
      const etat = viserJour(vise);
      if (!etat.ouvert) throw new Error('le calendrier ne s\'ouvre pas');
      if (!etat.vueMois) throw new Error('le calendrier n\'affiche pas les jours');
      if (etat.interdit) throw new Error('date non sélectionnable dans le calendrier');
      if (etat.cible) {
        clicSimule(etat.cible);
        await dormir(300);
        const valider = Array.from(document.querySelectorAll('.dx-overlay-wrapper .dx-button.dx-popup-done')).filter(visible).pop();
        if (valider) { clicSimule(valider); await dormir(200); }
        return;
      }
      if (!etat.mois) { await dormir(200); continue; }
      if (etat.mois === mois) throw new Error('jour absent du calendrier');
      clicSimule(etat.cal.querySelector(etat.mois < mois ? '.dx-calendar-navigator-next-view' : '.dx-calendar-navigator-previous-view'));
      for (let k = 0; k < 20; k++) { await dormir(150); if (viserJour(vise).mois !== etat.mois) break; }
      await dormir(250);
    }
    throw new Error('mois introuvable dans le calendrier');
  }
  const fermerCalendrier = () => { if (viserJour('').ouvert) clicSimule(document.body, true); };

  // Dernier recours : passer par l'objet DevExtreme, quand la page le laisse accessible.
  function parObjetDx(champ, jourIso) {
    const racine = champ.closest('.dx-datebox');
    let objet = null;
    try { objet = window.DevExpress.ui.dxDateBox.getInstance(racine); } catch (e) { /* non expose */ }
    if (!objet) try { objet = (window.jQuery || window.$)(racine).dxDateBox('instance'); } catch (e) { /* non expose */ }
    if (!objet) try { objet = window.ng.getComponent(racine.closest('dx-date-box') || racine).instance; } catch (e) { /* non expose */ }
    if (!objet || typeof objet.option !== 'function') return false;
    const actuelle = objet.option('value'), d = versDate(jourIso);
    let valeur = d;
    if (typeof actuelle === 'number') valeur = d.getTime();
    else if (typeof actuelle === 'string') {
      valeur = /Z$/.test(actuelle) ? d.toISOString().replace(/\.\d+Z$/, 'Z') : /T/.test(actuelle) ? jourIso + 'T00:00:00' : jourIso;
    }
    objet.option('value', valeur);
    return true;
  }

  let echecsSaisie = 0, methodeAnnoncee = false;
  async function reglerDate(champ, jourIso) {
    const dx = valeurDx(champ);
    if (dx === undefined) return reglerDateSimple(champ, jourIso);
    if (dx === jourIso) return;
    let obtenue = dx;
    for (const methode of ['saisie', 'calendrier', 'objet']) {
      try {
        if (methode === 'saisie') {
          const texte = echecsSaisie < 2 ? texteCommeAffiche(champ.value, valeurDx(champ), jourIso) : null;
          if (!texte) continue;
          saisieSimulee(champ, texte);
        } else if (methode === 'calendrier') await calendrierDx(champ, jourIso);
        else if (!parObjetDx(champ, jourIso)) continue;
      } catch (e) { journal('date par ' + methode + ' : ' + e.message); fermerCalendrier(); continue; }
      await calme();
      obtenue = valeurDx(champ);
      if (obtenue === jourIso) {
        if (!methodeAnnoncee) { methodeAnnoncee = true; journal('date réglée par ' + methode + ', valeur contrôlée'); }
        return;
      }
      if (methode === 'saisie') echecsSaisie++;
      fermerCalendrier();
    }
    throw new Error('date refusée par la page (le sélecteur indique ' + obtenue + ')');
  }

  // Champ date ordinaire (hors DevExtreme).
  async function reglerDateSimple(champ, jourIso) {
    const natifDate = (champ.getAttribute('type') || '').toLowerCase() === 'date';
    let texte = jourIso;
    if (!natifDate) {
      const m = /^(\d{1,2})([\/.\-])(\d{1,2})\2(\d{4})$/.exec(champ.value.trim());
      if (m && Number(m[3]) > 12) texte = [jourIso.slice(5, 7), jourIso.slice(8), jourIso.slice(0, 4)].join(m[2]);
      else if (!/^\d{4}-\d{2}-\d{2}$/.test(champ.value.trim())) texte = jourIso.split('-').reverse().join(m ? m[2] : '/');
    }
    if (champ.value.trim() === texte) return;
    saisieSimulee(champ, texte);
    await calme();
    if (champ.value.trim() !== texte) throw new Error('date refusée par la page (le champ affiche ' + champ.value + ')');
  }

  // Tableau exporte par le bouton : vide ?
  function tableauVide(bouton) {
    let grille = bouton.closest(GRILLES);
    if (!grille) {
      const toutes = Array.from(document.querySelectorAll(GRILLES)).filter(g => visible(g) && !g.parentElement.closest(GRILLES));
      if (toutes.length === 1) grille = toutes[0];
    }
    return !!grille && Array.from(grille.querySelectorAll('.dx-datagrid-nodata, .dx-treelist-nodata')).some(visible);
  }

  // ---------------------------------------------------------------- panneau
  const ui = {};
  const lignesJournal = [];
  function journal(message) {
    lignesJournal.push(new Date().toLocaleTimeString() + ' ' + message);
    if (lignesJournal.length > 400) lignesJournal.shift();
    try { console.log('[Mise au mille] ' + message); } catch (e) { /* console indisponible */ }
  }
  function creer(tag, style, texte) {
    const e = document.createElement(tag);
    if (style) e.style.cssText = style;
    if (texte) e.textContent = texte;
    return e;
  }
  const STYLE_BOUTON = 'margin:0 6px 0 0;font:bold 14px Arial,sans-serif;padding:6px 12px;cursor:pointer;border:0;border-radius:4px;';
  function bouton(parent, libelle, action, couleur) {
    const b = creer('button', STYLE_BOUTON + 'background:' + (couleur || '#eceff1') + ';color:' + (couleur ? '#fff' : '#000'), libelle);
    b.type = 'button';
    b.setAttribute('data-bot-bouton', '1');
    b.addEventListener('click', action);
    parent.appendChild(b);
    return b;
  }

  function construirePanneau() {
    const p = creer('div', 'position:fixed;left:0;right:0;bottom:0;z-index:2147483647;background:#1b2a33;color:#fff;' +
      'font:14px Arial,sans-serif;padding:10px 14px;box-shadow:0 -2px 10px rgba(0,0,0,.5);text-align:left;line-height:1.5');
    p.setAttribute('data-bot-barre', 'panneau');
    ui.panneau = p;
    ui.message = creer('div', 'font:bold 16px Arial,sans-serif;margin-bottom:6px');
    ui.detail = creer('div', 'margin-bottom:8px;color:#cfd8dc');
    ui.question = creer('div', 'margin-bottom:8px;display:none');
    ui.commandes = creer('div', '');
    p.append(ui.message, ui.detail, ui.question, ui.commandes);

    const STYLE_DATE = 'font:14px Arial,sans-serif;padding:4px;margin:0 10px 0 4px;border:0;border-radius:4px;color:#000;background:#fff';
    ui.commandes.appendChild(creer('span', '', 'Du'));
    ui.debut = creer('input', STYLE_DATE); ui.debut.type = 'date';
    ui.commandes.appendChild(ui.debut);
    ui.commandes.appendChild(creer('span', '', 'au'));
    ui.fin = creer('input', STYLE_DATE); ui.fin.type = 'date';
    ui.commandes.appendChild(ui.fin);
    let memo = {};
    try { memo = JSON.parse(localStorage.getItem('mam-reglages') || '{}'); } catch (e) { /* stockage local indisponible */ }
    ui.debut.value = memo.debut || REGLAGES.debut;
    ui.fin.value = isoLocal(new Date());

    ui.demarrer = bouton(ui.commandes, 'Démarrer', () => lancer(), '#2e7d32');
    ui.pause = bouton(ui.commandes, 'Pause', () => { etat.pauseDemandee = true; ui.pause.disabled = true; afficher('Pause demandée : arrêt à la fin du jour en cours...'); });
    ui.zip = bouton(ui.commandes, 'Télécharger le ZIP', () => telechargerZip(), '#1565c0');
    ui.effacer = bouton(ui.commandes, 'Tout effacer', () => toutEffacer());
    bouton(ui.commandes, 'Masquer', () => { p.style.display = 'none'; });
    ui.pause.style.display = 'none';
    document.documentElement.appendChild(p);
  }
  function afficher(message, detail) {
    ui.message.textContent = message;
    if (detail !== undefined) ui.detail.textContent = detail;
  }

  // Question posee dans le panneau : renvoie le rang du bouton choisi.
  function demander(message, libelles) {
    return new Promise(resolve => {
      ui.question.textContent = '';
      ui.question.style.display = '';
      ui.question.appendChild(creer('span', 'font-weight:bold;margin-right:10px;color:#ffeb3b', message));
      libelles.forEach((libelle, i) => bouton(ui.question, libelle, () => { ui.question.style.display = 'none'; resolve(i); }, i ? null : '#c62828'));
    });
  }

  // L'utilisateur montre un element en cliquant dessus (le clic n'atteint pas la page).
  function designer(message, optionnel, genre) {
    return new Promise(resolve => {
      ui.question.textContent = '';
      ui.question.style.display = '';
      const texte = creer('span', 'font-weight:bold;margin-right:10px;color:#ffeb3b', message);
      ui.question.appendChild(texte);
      const TYPES = ['click', 'mousedown', 'mouseup', 'pointerdown', 'pointerup', 'dblclick'];
      let survole = null, ancien = '';
      const quitterSurvol = () => { if (survole) survole.style.outline = ancien; survole = null; };
      const survol = e => {
        if (!horsPanneau(e.target) || e.target === survole) return;
        quitterSurvol();
        survole = e.target; ancien = survole.style.outline; survole.style.outline = '3px solid #ff1744';
      };
      const fin = v => {
        TYPES.forEach(t => document.removeEventListener(t, bloque, true));
        document.removeEventListener('mousemove', survol, true);
        quitterSurvol();
        ui.question.style.display = 'none';
        resolve(v);
      };
      const bloque = e => {
        if (!horsPanneau(e.target)) return;
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        if (e.type !== 'click') return;
        const cible = genre === 'date' ? versInput(e.target) : (e.target.closest(BOUTONS + ', [onclick]') || e.target);
        const r = cible && repere(cible);
        if (!r) { texte.textContent = 'Élément non reconnu, essaie encore. ' + message; return; }
        fin(r);
      };
      if (optionnel) bouton(ui.question, 'Il n\'y en a pas : passer', () => fin(null));
      bouton(ui.question, 'Abandonner', () => fin(undefined));
      TYPES.forEach(t => document.addEventListener(t, bloque, true));
      document.addEventListener('mousemove', survol, true);
    });
  }

  function entourer(el, contour) {
    const cible = el.closest('.dx-datebox') || el;
    if (contour) cible.scrollIntoView({block: 'center'});
    cible.style.outline = contour;
  }

  // Champ date, bouton d'extraction et bouton de validation eventuel.
  async function obtenirReperes() {
    let memo = null;
    try { memo = JSON.parse(localStorage.getItem('mam-reperes') || 'null'); } catch (e) { /* stockage local indisponible */ }
    if (memo && retrouver(memo.date) && retrouver(memo.exporter)) return memo;

    const champ = meilleur('input', scoreDate), boutonExport = meilleur(BOUTONS, scoreExport);
    let reperes = null;
    if (champ && boutonExport && repere(champ) && repere(boutonExport)) {
      entourer(champ, '4px solid #ff1744'); entourer(boutonExport, '4px solid #ff1744');
      afficher('Vérification des éléments de la page');
      const choix = await demander('J\'ai entouré en rouge le champ DATE et le bouton d\'EXTRACTION EXCEL. Est-ce correct ?',
                                   ['Oui, lancer', 'Non, je vais les montrer']);
      entourer(champ, ''); entourer(boutonExport, '');
      if (choix === 0) reperes = {date: repere(champ), exporter: repere(boutonExport), valider: null};
    }
    if (!reperes) {
      afficher('Montre-moi les éléments de la page (tes clics ne déclenchent rien)');
      const date = await designer('1/3 - Clique sur le CHAMP DATE.', false, 'date');
      if (!date) return null;
      const exporter = await designer('2/3 - Clique sur le BOUTON D\'EXTRACTION EXCEL.', false, 'bouton');
      if (!exporter) return null;
      const valider = await designer('3/3 - S\'il faut cliquer sur un bouton APRÈS avoir changé la date (Rechercher, Valider...), clique dessus.', true, 'bouton');
      if (valider === undefined) return null;
      reperes = {date, exporter, valider};
    }
    try { localStorage.setItem('mam-reperes', JSON.stringify(reperes)); } catch (e) { /* stockage local indisponible */ }
    return reperes;
  }

  // ------------------------------------------------------------- deroulement
  const etat = {enCours: false, pauseDemandee: false};
  const titreOrigine = document.title;

  async function bilan() {
    const fiches = await base.toutes();
    const compte = {ok: 0, vide: 0, erreur: 0, aucun: 0};
    fiches.forEach(f => { compte[f.statut] = (compte[f.statut] || 0) + 1; });
    return {fiches, compte, texte: compte.ok + ' fichiers en mémoire, ' + (compte.vide + compte.aucun) + ' jours sans fichier, ' + compte.erreur + ' en erreur'};
  }

  async function traiterJour(reperes, jourIso) {
    const champ = retrouver(reperes.date);
    if (!champ || !visible(champ)) throw new Error('champ date introuvable dans la page');
    await reglerDate(champ, jourIso);
    if (reperes.valider) {
      const valider = retrouver(reperes.valider);
      if (!valider) throw new Error('bouton de validation introuvable');
      reseau.marquer(); clicSimule(valider);
    }
    await calme();
    const boutonExport = retrouver(reperes.exporter);
    if (!boutonExport || !visible(boutonExport)) throw new Error('bouton d\'extraction introuvable dans la page');
    if (boutonExport.disabled || boutonExport.classList.contains('dx-state-disabled')) return {statut: 'aucun', remarque: 'bouton d\'extraction grisé'};
    if (tableauVide(boutonExport) && !REGLAGES.exporterSiVide) return {statut: 'vide', remarque: 'tableau vide ce jour-là'};
    const fichier = await exporter(boutonExport);
    if (!fichier) return {statut: 'aucun', remarque: 'aucun fichier produit par la page'};
    const extension = (/\.[a-z0-9]{2,5}$/i.exec(fichier.nom) || ['.xlsx'])[0];
    return {statut: 'ok', nom: REGLAGES.prefixe + jourIso + extension, blob: fichier.blob, taille: fichier.blob.size,
            empreinte: await empreinte(fichier.blob), remarque: 'nom d\'origine : ' + fichier.nom};
  }

  async function lancer() {
    if (etat.enCours) return;
    const debut = ui.debut.value, fin = ui.fin.value;
    if (!debut || !fin || debut > fin) { afficher('Dates invalides : choisis un début et une fin.'); return; }
    try { localStorage.setItem('mam-reglages', JSON.stringify({debut})); } catch (e) { /* stockage local indisponible */ }
    etat.enCours = true; etat.pauseDemandee = false;
    ui.demarrer.style.display = 'none';
    let arret = null;
    try {
      const deja = new Map((await base.toutes()).map(f => [f.date, f]));
      const jours = [];
      for (let d = versDate(debut); isoLocal(d) <= fin; d.setDate(d.getDate() + 1)) {
        const fiche = deja.get(isoLocal(d));
        if (!fiche || (fiche.statut !== 'ok' && fiche.statut !== 'vide')) jours.push(isoLocal(d));
      }
      if (!jours.length) { arret = 'Tout est déjà extrait pour cette période.'; return; }
      const reperes = await obtenirReperes();
      if (!reperes) { arret = 'Repérage abandonné.'; return; }
      ui.pause.style.display = ''; ui.pause.disabled = false;
      await calme();

      let erreursDeSuite = 0, faits = 0, precedente = null, tousIdentiques = true, nouveaux = [];
      const compte = {ok: 0, vide: 0, aucun: 0, erreur: 0};
      deja.forEach(f => { compte[f.statut] = (compte[f.statut] || 0) + 1; });
      const depart = performance.now();
      for (const jour of jours) {
        if (etat.pauseDemandee) { arret = 'En pause. Clique sur Démarrer pour reprendre.'; break; }
        const reste = faits ? Math.round((performance.now() - depart) / faits * (jours.length - faits) / 60000) : null;
        afficher('Extraction en cours : ' + jourFr(jour) + '  (' + (faits + 1) + '/' + jours.length + ')',
                 compte.ok + ' fichiers en mémoire, ' + (compte.vide + compte.aucun) + ' jours sans fichier, ' + compte.erreur + ' en erreur'
                 + (reste !== null ? ' - reste environ ' + reste + ' min' : '') + ' - garde cet onglet affiché');
        document.title = '[' + (faits + 1) + '/' + jours.length + '] ' + titreOrigine;
        let fiche;
        try {
          fiche = await traiterJour(reperes, jour);
          erreursDeSuite = 0;
        } catch (e) {
          fiche = {statut: 'erreur', remarque: e.message};
          erreursDeSuite++;
          fermerCalendrier();
        }
        if (fiche.statut === 'ok') {
          if (fiche.empreinte === precedente) fiche.remarque += ' ; CONTENU IDENTIQUE au fichier précédent';
          else if (precedente !== null) tousIdentiques = false;
          precedente = fiche.empreinte;
          nouveaux.push(jour);
        }
        journal(jourFr(jour) + ' : ' + fiche.statut + (fiche.nom ? ' ' + fiche.nom : '') + ' (' + fiche.remarque + ')');
        const ancienne = deja.get(jour);
        if (ancienne) compte[ancienne.statut]--;
        compte[fiche.statut]++;
        await base.ecrire({date: jour, horodatage: new Date().toISOString(), ...fiche});
        faits++;
        if (nouveaux.length >= REGLAGES.identiquesMax && tousIdentiques) {
          for (const j of nouveaux) await base.ecrire({date: j, statut: 'erreur', remarque: 'fichier écarté : identique aux autres'});
          arret = 'ARRÊT : les ' + nouveaux.length + ' premiers fichiers ont exactement le même contenu. La page ne tient pas ' +
                  'compte de la date. S\'il faut cliquer sur un bouton après avoir changé la date, clique sur « Tout effacer » ' +
                  'puis Démarrer, réponds « Non, je vais les montrer » et montre ce bouton à l\'étape 3/3.';
          break;
        }
        if (erreursDeSuite >= REGLAGES.erreursDeSuiteMax) {
          arret = 'ARRÊT : ' + erreursDeSuite + ' erreurs de suite (dernière : ' + fiche.remarque + '). Recharge la page (F5), relance le script, puis Démarrer : il reprendra aux jours manquants.';
          break;
        }
      }
      if (!arret) arret = 'TERMINÉ. Clique sur « Télécharger le ZIP » pour récupérer les fichiers.';
    } catch (e) {
      arret = 'ARRÊT : ' + e.message;
      journal('erreur : ' + (e.stack || e.message));
    } finally {
      etat.enCours = false;
      ui.demarrer.style.display = ''; ui.pause.style.display = 'none';
      document.title = titreOrigine;
      afficher(arret || 'Arrêté.', (await bilan()).texte);
    }
  }

  async function telechargerZip() {
    const {fiches, compte} = await bilan();
    if (!compte.ok) { afficher('Aucun fichier en mémoire pour le moment.'); return; }
    ui.zip.disabled = true;
    try {
      afficher('Préparation du ZIP...');
      const lignes = ['date;statut;fichier;taille_octets;remarque'].concat(fiches.map(f =>
        [f.date, f.statut, f.nom || '', f.taille || '', (f.remarque || '').replace(/[;\r\n]/g, ' ')].join(';')));
      const contenu = fiches.filter(f => f.statut === 'ok' && f.blob).map(f => ({nom: f.nom, blob: f.blob}));
      contenu.push({nom: 'rapport_extraction.csv', blob: new Blob(['﻿' + lignes.join('\r\n') + '\r\n'], {type: 'text/csv'})});
      const zip = await construireZip(contenu);
      const premiers = fiches.filter(f => f.statut === 'ok');
      const lien = document.createElement('a');
      lien.href = natif.creerUrl.call(URL, zip);
      lien.download = REGLAGES.prefixe + premiers[0].date + '_au_' + premiers[premiers.length - 1].date + '.zip';
      natif.clicLien.call(lien);
      setTimeout(() => URL.revokeObjectURL(lien.href), 120000);
      afficher('ZIP téléchargé : ' + compte.ok + ' fichiers. Regarde dans ton dossier Téléchargements.');
    } catch (e) {
      afficher('Impossible de préparer le ZIP : ' + e.message);
    } finally { ui.zip.disabled = false; }
  }

  async function toutEffacer() {
    if (etat.enCours) { afficher('Mets d\'abord en pause.'); return; }
    const {compte} = await bilan();
    if (await demander('Effacer les ' + compte.ok + ' fichiers en mémoire et les éléments mémorisés ?', ['Oui, tout effacer', 'Annuler']) !== 0) return;
    await base.vider();
    try { localStorage.removeItem('mam-reperes'); } catch (e) { /* stockage local indisponible */ }
    afficher('Mémoire effacée.', (await bilan()).texte);
  }

  construirePanneau();
  window.__mam = {
    version: VERSION, reglages: REGLAGES, journal: lignesJournal,
    montrer() { ui.panneau.style.display = ''; },
  };
  bilan().then(b => afficher('Extraction Mise au mille : choisis la période puis clique sur Démarrer.', b.texte));
})();
