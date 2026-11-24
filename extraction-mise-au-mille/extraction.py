#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Bot d'extraction des archives "Mise au mille".

Pour chaque jour entre la date de debut et aujourd'hui, le bot :
  1. regle la date dans la page,
  2. clique sur le bouton d'extraction Excel,
  3. enregistre le fichier sous un nom contenant la date du programme.

Il pilote le navigateur Edge deja installe sur Windows (via Playwright).
Les jours deja extraits sont sautes : on peut relancer le bot autant de fois
que necessaire, il reprend la ou il s'etait arrete.

Utilisation :  python extraction.py --help
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import time
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from urllib.parse import urlparse

try:
    from playwright.sync_api import Error as ErreurPlaywright
    from playwright.sync_api import sync_playwright
except ImportError:
    sys.exit(
        "Le module 'playwright' n'est pas installe.\n"
        "Lance d'abord 1_installer.bat (ou : python -m pip install playwright)."
    )

ICI = Path(__file__).resolve().parent
FICHIER_CONFIG = ICI / "config.json"
FICHIER_JOURNAL = ICI / "journal.log"
FICHIER_RAPPORT = ICI / "rapport_extraction.csv"
DOSSIER_DIAGNOSTIC = ICI / "diagnostic"
DOSSIER_ERREURS = ICI / "erreurs"

DEFAUTS = {
    # Adresse de la page (demandee au premier lancement si vide).
    "url": "",
    # Premier jour a extraire.
    "date_debut": "2025-07-01",
    # Dossier ou ranger les fichiers Excel (relatif a ce script, ou chemin complet).
    "dossier_sortie": "extractions",
    # Nom des fichiers : {date} est remplace par la date du programme.
    "nom_fichier": "Mise_au_mille_{date}",
    "format_date_fichier": "%Y-%m-%d",
    # Navigateur : "msedge" (Edge), "chrome" ou "chromium".
    "navigateur": "msedge",
    # Elements de la page (remplis automatiquement au premier lancement).
    "selecteur_date": "",
    "selecteur_date_2": "",  # 2e champ date eventuel (date de fin), mis a la meme date
    "selecteur_valider": "",  # bouton a cliquer apres avoir regle la date, s'il y en a un
    "selecteur_export": "",
    # Format de saisie de la date ("" = detection automatique, ex. "%d/%m/%Y").
    "format_date": "",
    # Touche pressee apres la saisie de la date ("" = aucune).
    "touche_validation": "Enter",
    # Pause (secondes) apres le chargement des donnees du jour.
    "pause_apres_date": 0.5,
    # Sans telechargement ni activite reseau pendant ce delai (s) : jour sans fichier.
    "delai_sans_reponse": 8,
    # Attente maximale (s) d'un telechargement.
    "attente_max": 180,
    # Requetes de fond a ignorer pour savoir si la page a fini de charger.
    "url_ignorees": r"signalr|/hubs?/|socket\.io|sockjs|negotiate|/poll",
    # Reglages avances (tests).
    "invisible": False,
    "chemin_navigateur": "",
}

ERREURS_DE_SUITE_MAX = 5
IDENTIQUES_MAX = 4


# --------------------------------------------------------------------------
# Script injecte dans la page : detection et designation des elements
# --------------------------------------------------------------------------
JS_BOT = r"""
(() => {
  if (window.__bot) return;
  const visible = el => {
    const r = el.getBoundingClientRect(), s = getComputedStyle(el);
    return r.width > 0 && r.height > 0 && s.visibility !== 'hidden' && s.display !== 'none';
  };
  // Les identifiants generes (mat-input-3, pr_id_12...) changent d'un chargement a l'autre.
  const idStable = id => !!id && !/^(mat|cdk|pr_id|ng|el-id|input-|v-|radix|headlessui|mui|:r)/i.test(id) &&
    !/\d{3,}/.test(id) && !/[-_:]\d+$/.test(id);
  const q = v => '"' + String(v).replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';
  const BOUTONS = 'button, a, [role=button], input[type=button], input[type=submit]';

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

  function texteCourt(el) {
    return (el.innerText || el.value || '').trim().replace(/\s+/g, ' ');
  }

  function decrire(el) {
    const tag = el.tagName.toLowerCase();
    const sels = [];
    if (idStable(el.id)) sels.push('#' + CSS.escape(el.id));
    for (const a of ['data-testid', 'data-test', 'data-cy', 'formcontrolname', 'name', 'aria-label',
                     'placeholder', 'title', 'mattooltip', 'ptooltip']) {
      const v = el.getAttribute(a);
      if (v) sels.push(tag + '[' + a + '=' + q(v) + ']');
    }
    if (tag === 'input') {
      if (/^(button|submit)$/i.test(el.type) && el.value) sels.push('input[value=' + q(el.value) + ']');
    } else {
      const t = texteCourt(el);
      if (t && t.length <= 50) { sels.push(tag + ':text-is(' + q(t) + ')'); sels.push(tag + ':has-text(' + q(t) + ')'); }
    }
    sels.push(chemin(el));
    if (el.id) sels.push('#' + CSS.escape(el.id));
    const attrs = ['type', 'id', 'name', 'placeholder', 'title', 'aria-label']
      .map(a => el.getAttribute(a) ? a + '=' + q(el.getAttribute(a)) : '').filter(Boolean).join(' ');
    return {
      selecteurs: sels,
      apercu: '<' + tag + (attrs ? ' ' + attrs : '') + '> ' + texteCourt(el).slice(0, 60),
    };
  }

  // ---- champ date --------------------------------------------------------
  function scoreDate(el) {
    if (!visible(el) || el.disabled) return 0;
    const type = (el.getAttribute('type') || 'text').toLowerCase();
    if (type === 'date') return 100;
    if (!/^(text|search|tel)$/.test(type)) return 0;
    let indices = Array.from(el.attributes).map(a => a.name + '=' + a.value).join(' ');
    for (let e = el.parentElement, i = 0; e && i < 3; e = e.parentElement, i++)
      indices += ' ' + e.tagName + ' ' + (e.getAttribute('class') || '');
    indices = indices.toLowerCase();
    if (/date-?picker|calendar|flatpickr|date-?input|date-?box|el-date|ant-picker|dp__|p-calendar/.test(indices)) return 90;
    if (/^\s*(\d{1,2}[\/.\-]\d{1,2}[\/.\-]\d{2,4}|\d{4}-\d{2}-\d{2})\s*$/.test(el.value || '')) return 80;
    const perso = ['placeholder', 'name', 'id', 'aria-label', 'formcontrolname']
      .map(a => el.getAttribute(a) || '').join(' ').toLowerCase();
    if (/\bdate|jj[\/.\-]mm|dd[\/.\-]mm|mm[\/.\-]dd/.test(perso)) return 60;
    return 0;
  }

  function trouverDate() {
    const xs = Array.from(document.querySelectorAll('input'))
      .map(el => [scoreDate(el), el]).filter(x => x[0] > 0).sort((a, b) => b[0] - a[0]);
    if (!xs.length || (xs.length > 1 && xs[1][0] === xs[0][0])) return null;
    return decrire(xs[0][1]);
  }

  // ---- bouton d'extraction ----------------------------------------------
  function indicesBouton(el) {
    const bits = [el.innerText, el.value, el.title, el.getAttribute('aria-label'), el.getAttribute('mattooltip'),
                  el.getAttribute('ptooltip'), el.getAttribute('class'), el.getAttribute('href'), el.getAttribute('download')];
    el.querySelectorAll('i, svg, img, span, mat-icon').forEach(c =>
      bits.push(c.getAttribute('class'), c.getAttribute('title'), c.getAttribute('alt')));
    return bits.filter(Boolean).join(' ').toLowerCase();
  }

  function scoreExport(el) {
    if (!visible(el)) return 0;
    const t = indicesBouton(el);
    let s = 0;
    if (/extraction|extraire|export/.test(t)) s += 50;
    if (/excel|xlsx?\b/.test(t)) s += 40;
    if (/t[ée]l[ée]charg|download|get_app/.test(t)) s += 30;
    if (s && /pdf|csv|import/.test(t)) s -= 30;
    return s;
  }

  function trouverExport() {
    const xs = Array.from(document.querySelectorAll(BOUTONS))
      .map(el => [scoreExport(el), el]).filter(x => x[0] > 0).sort((a, b) => b[0] - a[0]);
    if (!xs.length || (xs.length > 1 && xs[1][0] === xs[0][0])) return null;
    return decrire(xs[0][1]);
  }

  // ---- designation a la souris ------------------------------------------
  function versInput(el) {
    if (el.tagName === 'INPUT') return el;
    if (el.tagName === 'LABEL' && el.control && el.control.tagName === 'INPUT') return el.control;
    for (let e = el, i = 0; e && i < 4; e = e.parentElement, i++) {
      const ins = Array.from(e.querySelectorAll('input')).filter(visible);
      if (ins.length === 1) return ins[0];
      if (ins.length > 1) return null;
    }
    return null;
  }

  function choisir(message, optionnel, genre) {
    return new Promise(resolve => {
      const barre = document.createElement('div');
      barre.style.cssText = 'position:fixed;left:0;right:0;bottom:0;z-index:2147483647;background:#b71c1c;' +
        'color:#fff;font:bold 17px Arial,sans-serif;padding:14px;text-align:center;box-shadow:0 -2px 8px rgba(0,0,0,.5)';
      const texte = document.createElement('span');
      texte.textContent = message;
      barre.appendChild(texte);
      const ajouterBouton = (libelle, valeur) => {
        const b = document.createElement('button');
        b.textContent = libelle;
        b.style.cssText = 'margin-left:16px;font:bold 15px Arial,sans-serif;padding:6px 14px;cursor:pointer';
        b.addEventListener('click', () => fin(valeur));
        barre.appendChild(b);
      };
      if (optionnel) ajouterBouton("Il n'y en a pas : passer", null);
      ajouterBouton('Abandonner', {abandon: true});

      let survole = null, ancienContour = '';
      const TYPES = ['click', 'mousedown', 'mouseup', 'pointerdown', 'pointerup', 'dblclick'];
      const desurvoler = () => { if (survole) survole.style.outline = ancienContour; survole = null; };
      const survol = e => {
        if (barre.contains(e.target) || e.target === survole) return;
        desurvoler();
        survole = e.target; ancienContour = survole.style.outline; survole.style.outline = '3px solid #ff1744';
      };
      const fin = v => {
        TYPES.forEach(t => document.removeEventListener(t, bloque, true));
        document.removeEventListener('mousemove', survol, true);
        desurvoler();
        barre.remove();
        resolve(v);
      };
      const bloque = e => {
        if (barre.contains(e.target)) return;
        e.preventDefault(); e.stopPropagation(); e.stopImmediatePropagation();
        if (e.type !== 'click') return;
        const cible = genre === 'date' ? versInput(e.target) : (e.target.closest(BOUTONS + ', [onclick]') || e.target);
        if (!cible) { texte.textContent = "Ce n'est pas un champ de saisie, essaie encore. " + message; return; }
        fin(decrire(cible));
      };
      TYPES.forEach(t => document.addEventListener(t, bloque, true));
      document.addEventListener('mousemove', survol, true);
      document.documentElement.appendChild(barre);
    });
  }

  // ---- inventaire pour le diagnostic ------------------------------------
  function inventaire() {
    const lignes = [];
    document.querySelectorAll('input, select, textarea, ' + BOUTONS).forEach(el => {
      if (!visible(el)) return;
      const attrs = Array.from(el.attributes)
        .filter(a => !/^(style|value|d|src|href)$/.test(a.name))
        .map(a => a.name + '=' + q(a.value.slice(0, 80))).join(' ');
      lignes.push('<' + el.tagName.toLowerCase() + ' ' + attrs + '> texte=' + q(texteCourt(el).slice(0, 60)) +
        ' | ' + chemin(el));
    });
    return lignes;
  }

  window.__bot = {trouverDate, trouverExport, choisir, inventaire};
})();
"""


# --------------------------------------------------------------------------
# Petits utilitaires
# --------------------------------------------------------------------------
class ArretBot(Exception):
    """Erreur qui arrete tout le traitement (inutile de continuer)."""


class EchecJour(Exception):
    """Erreur limitee au jour en cours (on passe au suivant)."""


def journal(message: str) -> None:
    with FICHIER_JOURNAL.open("a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {message}\n")


def log(message: str = "") -> None:
    print(message, flush=True)
    journal(message)


def oui(question: str) -> bool:
    journal("QUESTION " + question)
    try:
        reponse = input(question + " [O/n] ").strip().lower()
    except EOFError:
        reponse = ""
    journal("REPONSE " + reponse)
    return reponse in ("", "o", "oui", "y", "yes")


def pause_entree(message: str) -> None:
    journal("PAUSE " + message)
    try:
        input(message)
    except EOFError:
        pass


def lire_jour(texte: str) -> date:
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(texte.strip(), fmt).date()
        except ValueError:
            pass
    raise argparse.ArgumentTypeError(f"date invalide : {texte!r} (attendu AAAA-MM-JJ ou JJ/MM/AAAA)")


def charger_config() -> dict:
    cfg = dict(DEFAUTS)
    if FICHIER_CONFIG.exists():
        try:
            cfg.update(json.loads(FICHIER_CONFIG.read_text(encoding="utf-8-sig")))
        except ValueError as e:
            sys.exit(f"config.json est illisible ({e}). Corrige-le ou supprime-le.")
    return cfg


def enregistrer_config(cfg: dict) -> None:
    FICHIER_CONFIG.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


# --------------------------------------------------------------------------
# Dates : format de saisie et relecture
# --------------------------------------------------------------------------
def detecter_format(valeur: str, placeholder: str) -> str:
    """Devine le format de saisie d'apres ce que le champ affiche."""
    v = (valeur or "").strip()
    m = re.fullmatch(r"(\d{1,2})([/.\-])(\d{1,2})\2(\d{4}|\d{2})", v)
    if m:
        a, sep, b = int(m[1]), m[2], int(m[3])
        an = "%Y" if len(m[4]) == 4 else "%y"
        francais, americain = f"%d{sep}%m{sep}{an}", f"%m{sep}%d{sep}{an}"
        auj = date.today()
        if b > 12 >= a:
            return americain
        if a <= 12 and b <= 12 and a != b and (a, b) == (auj.month, auj.day):
            return americain
        return francais
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return "%Y-%m-%d"
    p = (placeholder or "").strip().lower()
    if re.fullmatch(r"[jdmay/.\-]{6,10}", p):
        for source, cible in (("aaaa", "%Y"), ("yyyy", "%Y"), ("jj", "%d"), ("dd", "%d"),
                              ("mm", "%m"), ("aa", "%y"), ("yy", "%y")):
            p = p.replace(source, cible)
        if "%d" in p and "%m" in p and ("%Y" in p or "%y" in p):
            return p
    return "%d/%m/%Y"


def lire_date(valeur: str, fmt: str):
    """Relit la date affichee par le champ. None si elle n'est pas reconnaissable."""
    v = (valeur or "").strip()
    m = re.search(r"(\d{4})-(\d{1,2})-(\d{1,2})", v)
    if m:
        an, mois, jour = int(m[1]), int(m[2]), int(m[3])
    else:
        m = re.search(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4}|\d{2})(?!\d)", v)
        if not m:
            return None
        a, b, an = int(m[1]), int(m[2]), int(m[3])
        mois_avant = "%m" in fmt and "%d" in fmt and fmt.index("%m") < fmt.index("%d")
        jour, mois = (b, a) if mois_avant else (a, b)
        if an < 100:
            an += 2000
    try:
        return date(an, mois, jour)
    except ValueError:
        return None


# --------------------------------------------------------------------------
# Suivi du reseau et des telechargements
# --------------------------------------------------------------------------
class Reseau:
    """Suit les requetes de la page pour savoir quand elle a fini de charger."""

    SUIVIES = ("xhr", "fetch", "document")

    def __init__(self, contexte, ignorees: str):
        self.ignorees = re.compile(ignorees, re.I) if ignorees else None
        self.en_cours = {}
        self.derniere = time.monotonic()
        self.erreurs_http = []
        contexte.on("request", self._debut)
        contexte.on("requestfinished", self._fin)
        contexte.on("requestfailed", self._fin)
        contexte.on("response", self._reponse)

    def _debut(self, requete):
        if requete.resource_type not in self.SUIVIES:
            return
        if self.ignorees and self.ignorees.search(requete.url):
            return
        self.en_cours[requete] = True
        self.derniere = time.monotonic()
        journal(f"  > {requete.method} {requete.url}")

    def _fin(self, requete):
        if self.en_cours.pop(requete, None):
            self.derniere = time.monotonic()

    def _reponse(self, reponse):
        if reponse.status >= 400 and reponse.request in self.en_cours:
            self.erreurs_http.append(reponse.status)
            journal(f"  < HTTP {reponse.status} {reponse.url}")

    def oublier(self):
        """A appeler avant une action : ignore ce qui etait deja en cours."""
        self.en_cours.clear()
        self.erreurs_http.clear()
        self.derniere = time.monotonic()

    def calme_depuis(self) -> float:
        return 0.0 if self.en_cours else time.monotonic() - self.derniere


class Evenements:
    """Recupere les telechargements et les boites de dialogue de tous les onglets."""

    def __init__(self, contexte):
        self.telechargements = []
        self.dialogues = []
        contexte.on("page", self._brancher)
        for page in contexte.pages:
            self._brancher(page)

    def _brancher(self, page):
        page.on("download", self._telechargement)
        page.on("dialog", self._dialogue)

    def _telechargement(self, telechargement):
        self.telechargements.append(telechargement)

    def _dialogue(self, dialogue):
        self.dialogues.append(dialogue.message)
        journal(f"  boite de dialogue : {dialogue.message}")
        try:
            dialogue.accept()
        except ErreurPlaywright:
            pass

    def oublier(self):
        self.telechargements.clear()
        self.dialogues.clear()


def attendre_calme(page, reseau: Reseau, calme: float = 0.6, maxi: float = 30.0) -> None:
    debut = time.monotonic()
    page.wait_for_timeout(200)
    while time.monotonic() - debut < maxi:
        if reseau.calme_depuis() >= calme:
            return
        page.wait_for_timeout(100)


# --------------------------------------------------------------------------
# Navigateur
# --------------------------------------------------------------------------
def ouvrir_navigateur(p, cfg: dict):
    hote = urlparse(cfg["url"]).hostname or ""
    arguments = ["--start-maximized"]
    if hote:
        # Connexion automatique avec la session Windows (SSO) pour ce serveur.
        arguments.append(f"--auth-server-allowlist={hote}")
    canaux = [cfg["navigateur"]] + [c for c in ("msedge", "chrome", "chromium") if c != cfg["navigateur"]]
    echecs = []
    for canal in canaux:
        options = dict(
            user_data_dir=str(ICI / "profil_navigateur" / canal),
            headless=bool(cfg["invisible"]),
            accept_downloads=True,
            no_viewport=True,
            args=arguments,
        )
        if cfg["chemin_navigateur"]:
            options["executable_path"] = cfg["chemin_navigateur"]
        elif canal != "chromium":
            options["channel"] = canal
        try:
            contexte = p.chromium.launch_persistent_context(**options)
            journal(f"navigateur ouvert : {canal}")
            return contexte
        except ErreurPlaywright as e:
            echecs.append(f"  - {canal} : {str(e).splitlines()[0]}")
    raise ArretBot(
        "Impossible d'ouvrir un navigateur :\n" + "\n".join(echecs) + "\n"
        "Si Edge est bloque par la politique de l'entreprise, essaie :\n"
        "  python -m playwright install chromium\n"
        'puis mets "navigateur": "chromium" dans config.json.'
    )


def aller_sur_la_page(page, cfg: dict) -> None:
    page.goto("about:blank")
    page.goto(cfg["url"], wait_until="domcontentloaded", timeout=60000)


def est_visible(page, selecteur: str, delai_ms: int) -> bool:
    if not selecteur:
        return False
    try:
        page.locator(selecteur).first.wait_for(state="visible", timeout=delai_ms)
        return True
    except ErreurPlaywright:
        return False


# --------------------------------------------------------------------------
# Reperage des elements de la page (automatique, sinon a la souris)
# --------------------------------------------------------------------------
def selecteur_unique(page, description):
    """Premier selecteur propose qui designe exactement un element."""
    if not description:
        return None
    for selecteur in description["selecteurs"]:
        try:
            if page.locator(selecteur).count() == 1:
                return selecteur
        except ErreurPlaywright:
            continue
    return None


def entourer(page, selecteur: str, contour: str) -> None:
    try:
        page.locator(selecteur).first.evaluate(
            "(e, c) => { if (c) e.scrollIntoView({block: 'center'}); e.style.outline = c; }", contour
        )
    except ErreurPlaywright:
        pass


def designer(page, message: str, optionnel: bool, genre: str):
    """Demande a l'utilisateur de cliquer sur un element ; renvoie son selecteur."""
    while True:
        page.evaluate(JS_BOT)
        description = page.evaluate(
            "([m, o, g]) => window.__bot.choisir(m, o, g)", [message, optionnel, genre]
        )
        if description is None:
            return ""
        if description.get("abandon"):
            raise ArretBot("Reperage abandonne.")
        selecteur = selecteur_unique(page, description)
        if selecteur:
            log(f"   -> {description['apercu']}")
            journal(f"   selecteur : {selecteur}")
            return selecteur
        log("   Je n'arrive pas a identifier cet element de facon sure, recommence.")


def reperer_elements(page, cfg: dict) -> None:
    """Trouve le champ date et le bouton d'extraction, puis les memorise."""
    log("Recherche du champ date et du bouton d'extraction dans la page...")
    page.evaluate(JS_BOT)
    champ = page.evaluate("() => window.__bot.trouverDate()")
    bouton = page.evaluate("() => window.__bot.trouverExport()")
    sel_date, sel_export, sel_valider = selecteur_unique(page, champ), selecteur_unique(page, bouton), ""

    if sel_date and sel_export:
        entourer(page, sel_date, "4px solid #ff1744")
        entourer(page, sel_export, "4px solid #ff1744")
        log(f"  Champ date          : {champ['apercu']}")
        log(f"  Bouton d'extraction : {bouton['apercu']}")
        log("Ces deux elements sont entoures en ROUGE dans la fenetre du navigateur.")
        correct = oui("Est-ce bien le champ date et le bouton d'extraction Excel ?")
        entourer(page, sel_date, "")
        entourer(page, sel_export, "")
        if not correct:
            sel_date = sel_export = None

    if not (sel_date and sel_export):
        log("")
        log("Montre-moi les elements : regarde la fenetre du navigateur et suis les")
        log("consignes du bandeau rouge en bas de la page (rien ne sera declenche).")
        sel_date = designer(page, "1/3 - Clique sur le CHAMP DATE (la case ou la date est affichee).", False, "date")
        sel_export = designer(page, "2/3 - Clique sur le BOUTON D'EXTRACTION EXCEL (aucune extraction ne sera lancee).",
                              False, "bouton")
        sel_valider = designer(
            page, "3/3 - S'il faut cliquer sur un bouton APRES avoir change la date (Rechercher, Valider...), clique dessus.",
            True, "bouton")

    cfg.update(selecteur_date=sel_date, selecteur_export=sel_export, selecteur_valider=sel_valider)
    enregistrer_config(cfg)
    log("Elements memorises dans config.json (relance avec --reapprendre pour les changer).")


def diagnostic(page, raison: str) -> None:
    """Enregistre de quoi comprendre la page quand le bot n'y arrive pas."""
    try:
        DOSSIER_DIAGNOSTIC.mkdir(exist_ok=True)
        page.evaluate(JS_BOT)
        lignes = [f"Raison : {raison}", f"URL : {page.url}", f"Titre : {page.title()}",
                  f"Cadres : {[c.url for c in page.frames]}", ""]
        lignes += page.evaluate("() => window.__bot.inventaire()")
        (DOSSIER_DIAGNOSTIC / "elements.txt").write_text("\n".join(lignes), encoding="utf-8")
        (DOSSIER_DIAGNOSTIC / "page.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(DOSSIER_DIAGNOSTIC / "page.png"), full_page=True)
        log(f"Diagnostic enregistre dans : {DOSSIER_DIAGNOSTIC}")
        log("  -> elements.txt (liste des champs et boutons) suffit en general pour ajuster le bot.")
    except Exception as e:  # le diagnostic ne doit jamais faire planter le bot
        log(f"(diagnostic impossible : {e})")


# --------------------------------------------------------------------------
# Reglage de la date
# --------------------------------------------------------------------------
def saisie_clavier(page, champ, texte: str, cfg: dict) -> None:
    champ.focus()
    champ.select_text()
    champ.press_sequentially(texte, delay=20)
    if cfg["touche_validation"]:
        page.keyboard.press(cfg["touche_validation"])
    champ.evaluate("e => e.blur()")


def saisie_directe(page, champ, texte: str, cfg: dict) -> None:
    champ.fill(texte)
    champ.dispatch_event("change")
    champ.evaluate("e => e.blur()")


def saisie_js(page, champ, texte: str, cfg: dict) -> None:
    champ.evaluate(
        """(e, v) => {
            Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set.call(e, v);
            for (const t of ['input', 'change', 'blur']) e.dispatchEvent(new Event(t, {bubbles: true}));
        }""",
        texte,
    )


METHODES = [("clavier", saisie_clavier), ("saisie directe", saisie_directe), ("javascript", saisie_js)]


class Saisie:
    """Regle la date dans un champ et verifie que la page l'a bien prise."""

    def __init__(self, cfg: dict):
        self.cfg = cfg
        self.format = cfg["format_date"]
        self.methode = 0  # indice de la methode qui marche, memorise d'un jour a l'autre
        self.non_verifiable = False

    def regler(self, page, reseau: Reseau, selecteur: str, jour: date) -> None:
        champ = page.locator(selecteur).first
        if not est_visible(page, selecteur, 20000):
            raise EchecJour("champ date introuvable dans la page")
        if (champ.get_attribute("type") or "").lower() == "date":
            reseau.oublier()
            champ.fill(jour.isoformat())
            champ.evaluate("e => e.blur()")
            attendre_calme(page, reseau)
            if champ.input_value() != jour.isoformat():
                raise EchecJour(f"le champ date affiche {champ.input_value()!r}")
            return

        if not self.format:
            self.format = detecter_format(champ.input_value(), champ.get_attribute("placeholder") or "")
            log(f"  Format de saisie de la date : {jour.strftime(self.format)} ({self.format})")
        texte = jour.strftime(self.format)

        affiche = ""
        for indice in range(self.methode, len(METHODES)):
            nom, methode = METHODES[indice]
            avant = champ.input_value()
            reseau.oublier()
            try:
                methode(page, champ, texte, self.cfg)
            except ErreurPlaywright as e:
                journal(f"  methode {nom} : {str(e).splitlines()[0]}")
                continue
            attendre_calme(page, reseau)
            affiche = champ.input_value()
            relu = lire_date(affiche, self.format)
            verifie = relu == jour and nom != "javascript"
            plausible = (relu == jour) or (relu is None and affiche.strip() and affiche != avant)
            if verifie or plausible:
                if indice != self.methode:
                    log(f"  (date reglee par la methode : {nom})")
                    self.methode = indice
                if not verifie and not self.non_verifiable:
                    self.non_verifiable = True
                    log("  ATTENTION : impossible de verifier que la page a bien pris la date."
                        " Controle les premiers fichiers extraits.")
                return
        raise EchecJour(f"date refusee par la page (le champ affiche {affiche!r} au lieu de {texte!r})")


# --------------------------------------------------------------------------
# Extraction d'un jour
# --------------------------------------------------------------------------
def empreinte(fichier: Path) -> str:
    """Empreinte du contenu, sans les metadonnees (dates de creation) d'un .xlsx."""
    h = hashlib.sha256()
    try:
        with zipfile.ZipFile(fichier) as z:
            for nom in sorted(z.namelist()):
                if not nom.startswith("docProps/"):
                    h.update(nom.encode())
                    h.update(z.read(nom))
    except zipfile.BadZipFile:
        h.update(fichier.read_bytes())
    return h.hexdigest()


def cliquer(page, bouton) -> None:
    try:
        bouton.click(timeout=8000)
    except ErreurPlaywright:
        # Un calendrier ou un menu resté ouvert peut masquer le bouton.
        page.keyboard.press("Escape")
        bouton.click(timeout=8000)


def extraire_jour(page, cfg: dict, reseau: Reseau, evenements: Evenements, saisie: Saisie,
                  jour: date, cible_sans_extension: Path):
    """Renvoie (statut, fichier ou None, remarque)."""
    evenements.oublier()
    saisie.regler(page, reseau, cfg["selecteur_date"], jour)
    if cfg["selecteur_date_2"]:
        saisie.regler(page, reseau, cfg["selecteur_date_2"], jour)
    if cfg["selecteur_valider"]:
        reseau.oublier()
        cliquer(page, page.locator(cfg["selecteur_valider"]).first)
        attendre_calme(page, reseau)
    page.wait_for_timeout(int(float(cfg["pause_apres_date"]) * 1000))

    bouton = page.locator(cfg["selecteur_export"]).first
    if not est_visible(page, cfg["selecteur_export"], 15000):
        raise EchecJour("bouton d'extraction introuvable dans la page")
    if not bouton.is_enabled():
        page.wait_for_timeout(2000)
        if not bouton.is_enabled():
            return "AUCUN FICHIER", None, "bouton d'extraction grise"

    evenements.oublier()
    reseau.oublier()
    cliquer(page, bouton)

    debut = time.monotonic()
    silence = float(cfg["delai_sans_reponse"])
    while not evenements.telechargements:
        ecoule = time.monotonic() - debut
        if ecoule > float(cfg["attente_max"]):
            raise EchecJour(f"pas de fichier apres {int(ecoule)} s (la page charge encore)")
        if ecoule >= silence and reseau.calme_depuis() >= silence:
            remarque = "; ".join(evenements.dialogues)
            if reseau.erreurs_http:
                codes = ", ".join(str(c) for c in sorted(set(reseau.erreurs_http)))
                if any(c in (401, 403) or c >= 500 for c in reseau.erreurs_http):
                    raise EchecJour(f"le serveur a repondu HTTP {codes}")
                remarque = (remarque + " " if remarque else "") + f"reponse HTTP {codes}"
            return "AUCUN FICHIER", None, remarque or "aucun telechargement declenche"
        page.wait_for_timeout(150)

    telechargement = evenements.telechargements[0]
    echec = telechargement.failure()
    if echec:
        raise EchecJour(f"telechargement en echec : {echec}")
    extension = Path(telechargement.suggested_filename).suffix or ".xlsx"
    cible = cible_sans_extension.with_name(cible_sans_extension.name + extension)
    telechargement.save_as(str(cible))
    if cible.stat().st_size == 0:
        cible.unlink()
        raise EchecJour("fichier telecharge vide (0 octet)")
    return "OK", cible, f"nom d'origine : {telechargement.suggested_filename}"


def fermer_onglets_en_trop(contexte, page) -> None:
    for autre in contexte.pages:
        if autre is not page:
            try:
                autre.close()
            except ErreurPlaywright:
                pass


# --------------------------------------------------------------------------
# Programme principal
# --------------------------------------------------------------------------
def lire_arguments():
    parser = argparse.ArgumentParser(description="Extraction des archives Mise au mille (un Excel par jour).")
    parser.add_argument("--debut", type=lire_jour, help="premier jour (AAAA-MM-JJ ou JJ/MM/AAAA)")
    parser.add_argument("--fin", type=lire_jour, help="dernier jour (par defaut : aujourd'hui)")
    parser.add_argument("--test", action="store_true", help="essai sur les 5 derniers jours seulement")
    parser.add_argument("--reapprendre", action="store_true",
                        help="reperer a nouveau le champ date et le bouton d'extraction")
    parser.add_argument("--diagnostic", action="store_true",
                        help="enregistrer une description de la page, sans rien extraire")
    parser.add_argument("--accepter-identiques", action="store_true",
                        help="ne pas s'arreter si les premiers fichiers ont tous le meme contenu")
    return parser.parse_args()


def fichier_existant(dossier: Path, nom: str):
    for f in dossier.glob(nom + ".*"):
        if f.suffix.lower() not in (".crdownload", ".tmp", ".part") and f.stat().st_size > 0:
            return f
    return None


def executer(args, cfg: dict) -> int:
    fin = args.fin or date.today()
    debut = args.debut or lire_jour(cfg["date_debut"])
    if args.test:
        debut = fin - timedelta(days=4)
    if debut > fin:
        sys.exit(f"La date de debut ({debut}) est apres la date de fin ({fin}).")
    jours = [debut + timedelta(days=i) for i in range((fin - debut).days + 1)]

    sortie = Path(cfg["dossier_sortie"])
    if not sortie.is_absolute():
        sortie = ICI / sortie
    sortie.mkdir(parents=True, exist_ok=True)

    def nom_du_jour(j: date) -> str:
        return cfg["nom_fichier"].format(date=j.strftime(cfg["format_date_fichier"]))

    a_faire = [j for j in jours if not fichier_existant(sortie, nom_du_jour(j))]
    log(f"Periode : du {debut:%d/%m/%Y} au {fin:%d/%m/%Y} ({len(jours)} jours)")
    log(f"Deja extraits : {len(jours) - len(a_faire)}  |  A traiter : {len(a_faire)}")
    log(f"Dossier des fichiers : {sortie}")
    if not a_faire and not args.diagnostic:
        log("Tout est deja extrait, rien a faire.")
        return 0

    compte = {"OK": 0, "AUCUN FICHIER": 0, "ERREUR": 0}
    sans_fichier, en_erreur, identiques, extraits = [], [], [], []

    with sync_playwright() as p:
        contexte = ouvrir_navigateur(p, cfg)
        try:
            contexte.add_init_script(JS_BOT)
            reseau = Reseau(contexte, cfg["url_ignorees"])
            evenements = Evenements(contexte)
            page = contexte.pages[0] if contexte.pages else contexte.new_page()
            page.set_default_timeout(30000)
            log("Ouverture de la page...")
            try:
                aller_sur_la_page(page, cfg)
            except ErreurPlaywright as e:
                raise ArretBot(f"Impossible d'ouvrir {cfg['url']} : {str(e).splitlines()[0]}\n"
                               "Verifie que tu es bien sur le reseau de l'usine (ou le VPN).")

            if args.reapprendre:
                cfg.update(selecteur_date="", selecteur_export="", selecteur_valider="")
            connu = cfg["selecteur_date"] and cfg["selecteur_export"]
            if not (connu and est_visible(page, cfg["selecteur_date"], 15000)):
                log("")
                log("Dans la fenetre du navigateur : connecte-toi si c'est demande et verifie")
                log("que la page 'Mise au mille' est bien affichee.")
                pause_entree("Appuie ensuite sur ENTREE ici pour continuer... ")
                if args.diagnostic:
                    diagnostic(page, "demande par --diagnostic")
                    return 0
                if not (connu and est_visible(page, cfg["selecteur_date"], 3000)):
                    try:
                        reperer_elements(page, cfg)
                    except ArretBot:
                        diagnostic(page, "reperage des elements impossible")
                        raise
            elif args.diagnostic:
                diagnostic(page, "demande par --diagnostic")
                return 0

            nouveau = not FICHIER_RAPPORT.exists()
            with FICHIER_RAPPORT.open("a", newline="", encoding="utf-8-sig") as f_rapport:
                rapport = csv.writer(f_rapport, delimiter=";")
                if nouveau:
                    rapport.writerow(["horodatage", "date", "statut", "fichier", "taille_octets", "remarque"])

                saisie = Saisie(cfg)
                erreurs_de_suite = 0
                derniere_empreinte, tous_identiques = None, True
                log("")
                for numero, jour in enumerate(a_faire, 1):
                    prefixe = f"[{numero}/{len(a_faire)}] {jour:%d/%m/%Y}"
                    fichier = None
                    try:
                        statut, fichier, remarque = extraire_jour(
                            page, cfg, reseau, evenements, saisie, jour, sortie / nom_du_jour(jour))
                    except (EchecJour, ErreurPlaywright, OSError) as e:
                        statut, remarque = "ERREUR", str(e).splitlines()[0]
                        if "closed" in remarque.lower():
                            raise ArretBot("La fenetre du navigateur a ete fermee.")
                    fermer_onglets_en_trop(contexte, page)

                    taille = ""
                    if statut == "OK":
                        extraits.append(fichier)
                        taille = fichier.stat().st_size
                        e = empreinte(fichier)
                        if e == derniere_empreinte:
                            identiques.append(jour)
                            remarque += " ; CONTENU IDENTIQUE au fichier precedent"
                        elif derniere_empreinte is not None:
                            tous_identiques = False
                        derniere_empreinte = e
                        erreurs_de_suite = 0
                        log(f"{prefixe}  OK   {fichier.name} ({taille // 1024 or 1} Ko)"
                            + ("  [identique au precedent]" if identiques and identiques[-1] == jour else ""))
                    elif statut == "AUCUN FICHIER":
                        sans_fichier.append(jour)
                        erreurs_de_suite = 0
                        log(f"{prefixe}  --   aucun fichier ({remarque})")
                    else:
                        en_erreur.append(jour)
                        erreurs_de_suite += 1
                        log(f"{prefixe}  ERREUR  {remarque}")
                        try:
                            DOSSIER_ERREURS.mkdir(exist_ok=True)
                            page.screenshot(path=str(DOSSIER_ERREURS / f"{jour.isoformat()}.png"))
                        except ErreurPlaywright:
                            pass
                    compte[statut] += 1
                    rapport.writerow([f"{datetime.now():%Y-%m-%d %H:%M:%S}", jour.isoformat(), statut,
                                      fichier.name if fichier else "", taille, remarque])
                    f_rapport.flush()

                    if (compte["OK"] >= IDENTIQUES_MAX and tous_identiques and not args.accepter_identiques):
                        # Fichiers probablement mal dates : on les met de cote pour qu'ils ne
                        # passent pas pour des jours deja extraits.
                        suspects = sortie / "_suspects"
                        suspects.mkdir(exist_ok=True)
                        for f in extraits:
                            f.replace(suspects / f.name)
                        diagnostic(page, "les premiers fichiers extraits sont tous identiques")
                        raise ArretBot(
                            f"Les {compte['OK']} premiers fichiers ont exactement le meme contenu : la page ne\n"
                            "semble pas tenir compte de la date que je saisis. Je les ai deplaces dans\n"
                            f"{suspects}\n"
                            "Ouvre-les pour verifier. S'ils sont normaux (jours reellement identiques),\n"
                            "relance avec --accepter-identiques. Sinon le bot doit etre ajuste a cette\n"
                            "page : envoie le fichier diagnostic/elements.txt.")
                    if erreurs_de_suite >= ERREURS_DE_SUITE_MAX:
                        diagnostic(page, f"{erreurs_de_suite} erreurs de suite, derniere : {remarque}")
                        raise ArretBot(
                            f"{erreurs_de_suite} erreurs de suite, j'arrete (session expiree ? page changee ?).\n"
                            "Relance le bot : il reprendra aux jours manquants. Si l'erreur revient des le\n"
                            "debut, relance avec --reapprendre pour lui remontrer le champ date et le bouton.")
                    if statut == "ERREUR":
                        try:  # on repart d'une page propre
                            aller_sur_la_page(page, cfg)
                            est_visible(page, cfg["selecteur_date"], 20000)
                        except ErreurPlaywright:
                            pass
        except KeyboardInterrupt:
            log("\nInterrompu (Ctrl+C).")
        except ArretBot as e:
            log("")
            log(f"ARRET : {e}")
        finally:
            try:
                contexte.close()
            except Exception:
                pass

    log("")
    log("=" * 60)
    log(f"Fichiers extraits : {compte['OK']}   |   Jours sans fichier : {compte['AUCUN FICHIER']}"
        f"   |   Erreurs : {compte['ERREUR']}")
    restants = len(a_faire) - sum(compte.values())
    if restants:
        log(f"Jours non traites : {restants} (relance le bot pour continuer)")
    if sans_fichier:
        log("Jours sans fichier (week-ends, arrets... a verifier) : "
            + ", ".join(f"{j:%d/%m/%Y}" for j in sans_fichier[:40]) + (" ..." if len(sans_fichier) > 40 else ""))
    if identiques:
        log(f"Fichiers au contenu identique au jour precedent : {len(identiques)} (voir le rapport)")
    if en_erreur:
        log("Jours en erreur (seront retentes au prochain lancement) : "
            + ", ".join(f"{j:%d/%m/%Y}" for j in en_erreur[:40]) + (" ..." if len(en_erreur) > 40 else ""))
    log(f"Rapport detaille : {FICHIER_RAPPORT}")
    return 1 if (en_erreur or restants) else 0


def main() -> int:
    for flux in (sys.stdout, sys.stderr):
        try:
            flux.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    args = lire_arguments()
    cfg = charger_config()
    journal("=" * 30 + " lancement : " + " ".join(sys.argv[1:]))
    if not cfg["url"]:
        cfg["url"] = input("Adresse de la page 'Mise au mille' (copie-la depuis le navigateur) : ").strip()
        if not cfg["url"]:
            sys.exit("Aucune adresse fournie.")
        enregistrer_config(cfg)
    try:
        return executer(args, cfg)
    except ArretBot as e:
        log(f"ARRET : {e}")
        return 2
    except PermissionError as e:
        log(f"ARRET : fichier verrouille ({e.filename}). Ferme-le (ouvert dans Excel ?) et relance.")
        return 2


if __name__ == "__main__":
    sys.exit(main())
