#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Assemble les fichiers Excel journaliers "Mise au mille" en un seul fichier.

A partir du ZIP produit par l'extraction (ou d'un dossier de fichiers .xlsx) :
  - un fichier unique, avec une seule ligne de titres, toutes les lignes de tous
    les jours a la suite, et une premiere colonne "Date programme" ;
  - une copie de chaque fichier journalier avec cette meme colonne ajoutee.

La date vient du nom de chaque fichier (Mise_au_mille_2025-07-01.xlsx).
Aucun module a installer : Python seul suffit.

Utilisation :  python assembler.py [fichier.zip | dossier | fichiers .xlsx] [options]
"""
from __future__ import annotations

import argparse
import io
import os
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

ICI = Path(__file__).resolve().parent
TITRE_DATE = "Date programme"
MOTIF_DATE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")
MOTIF_ZIP = "Mise_au_mille_*_au_*.zip"
LIGNES_MAX = 1048576
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Formats de nombre integres d'Excel qui affichent une date ou une heure.
FORMATS_DATE_INTEGRES = set(range(14, 23)) | set(range(27, 37)) | set(range(45, 48)) | set(range(50, 59))


class Arret(Exception):
    """Erreur expliquee a l'utilisateur, sans trace technique."""


# --------------------------------------------------------------------------
# Lecture d'un classeur .xlsx (bibliotheque standard uniquement)
# --------------------------------------------------------------------------
def sans_espace_de_noms(racine):
    for element in racine.iter():
        if isinstance(element.tag, str) and "}" in element.tag:
            element.tag = element.tag.split("}", 1)[1]
    return racine


def indice_colonne(reference: str) -> int:
    """'A1' -> 0, 'AB12' -> 27."""
    n = 0
    for caractere in reference:
        if not caractere.isalpha():
            break
        n = n * 26 + (ord(caractere.upper()) - 64)
    return n - 1


def lettre_colonne(indice: int) -> str:
    lettres = ""
    indice += 1
    while indice:
        indice, reste = divmod(indice - 1, 26)
        lettres = chr(65 + reste) + lettres
    return lettres


def texte_riche(element) -> str:
    """Texte d'un <si> ou <is> : texte direct ou morceaux mis en forme."""
    if element is None:
        return ""
    morceaux = []
    for enfant in element:
        if enfant.tag == "t":
            morceaux.append(enfant.text or "")
        elif enfant.tag == "r":
            morceaux.extend(t.text or "" for t in enfant.findall("t"))
    return "".join(morceaux)


def est_format_date(cle) -> bool:
    if cle is None:
        return False
    if isinstance(cle, int):
        return cle in FORMATS_DATE_INTEGRES
    code = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.', "", cle).lower()
    return bool(re.search(r"[dmyhs]", code))


class Classeur:
    """Premiere feuille d'un .xlsx : lignes de cellules (genre, valeur, format)."""

    def __init__(self, octets: bytes):
        try:
            self.zip = zipfile.ZipFile(io.BytesIO(octets))
        except zipfile.BadZipFile:
            raise Arret("ce n'est pas un fichier .xlsx valide")
        self.noms = {n.lower(): n for n in self.zip.namelist() if not n.endswith("/")}
        chemin, self.dates_1904 = self._localiser_feuille()
        chaines = self._lire_chaines()
        formats = self._lire_formats()
        feuille = self._xml(chemin)
        self.largeurs = [(int(c.get("min", "1")), int(c.get("max", "1")), c.get("width"))
                         for c in feuille.findall("cols/col") if c.get("width")]
        self.lignes = []
        for rangee in feuille.iter("row"):
            cellules, colonne = {}, -1
            for c in rangee.findall("c"):
                reference = c.get("r")
                colonne = indice_colonne(reference) if reference else colonne + 1
                genre = c.get("t", "n")
                v = c.find("v")
                brut = v.text if v is not None and v.text is not None else ""
                if genre == "s":
                    cellule = ("s", chaines[int(brut)] if brut.strip() else "", None)
                elif genre == "inlineStr":
                    cellule = ("s", texte_riche(c.find("is")), None)
                elif genre in ("str", "e", "d"):
                    cellule = ("s", brut, None)
                elif genre == "b":
                    cellule = ("b", brut.strip(), None)
                else:
                    style = int(c.get("s") or 0)
                    cle = formats[style] if style < len(formats) else None
                    if self.dates_1904 and brut.strip() and est_format_date(cle):
                        brut = repr(float(brut) + 1462)
                    cellule = ("n", brut.strip(), cle)
                    if cellule[1]:
                        try:
                            float(cellule[1])
                        except ValueError:
                            cellule = ("s", brut, None)
                if cellule[1] != "":
                    cellules[colonne] = cellule
            if cellules:
                ligne = [None] * (max(cellules) + 1)
                for colonne, cellule in cellules.items():
                    ligne[colonne] = cellule
                self.lignes.append(ligne)

    def _xml(self, nom: str):
        return sans_espace_de_noms(ET.fromstring(self.zip.read(self.noms[nom.lower()])))

    def _localiser_feuille(self):
        cible, dates_1904 = None, False
        if "xl/workbook.xml" in self.noms:
            classeur = self._xml("xl/workbook.xml")
            proprietes = classeur.find("workbookPr")
            dates_1904 = proprietes is not None and proprietes.get("date1904") in ("1", "true")
            feuilles = classeur.findall("sheets/sheet")
            lien = next((v for k, v in feuilles[0].attrib.items() if k.endswith("}id")), None) if feuilles else None
            if lien and "xl/_rels/workbook.xml.rels" in self.noms:
                for relation in self._xml("xl/_rels/workbook.xml.rels"):
                    if relation.get("Id") == lien:
                        chemin = relation.get("Target", "")
                        cible = chemin.lstrip("/") if chemin.startswith("/") else posixpath.normpath("xl/" + chemin)
        if not cible or cible.lower() not in self.noms:
            feuilles = sorted(n for n in self.noms if re.fullmatch(r"xl/worksheets/[^/]+\.xml", n))
            if not feuilles:
                raise Arret("aucune feuille de calcul dans ce fichier")
            cible = feuilles[0]
        return cible, dates_1904

    def _lire_chaines(self):
        if "xl/sharedstrings.xml" not in self.noms:
            return []
        return [texte_riche(si) for si in self._xml("xl/sharedStrings.xml").findall("si")]

    def _lire_formats(self):
        """Pour chaque style de cellule : None (standard), numero de format integre, ou code de format."""
        if "xl/styles.xml" not in self.noms:
            return []
        styles = self._xml("xl/styles.xml")
        personnalises = {int(f.get("numFmtId")): f.get("formatCode", "") for f in styles.findall("numFmts/numFmt")}
        formats = []
        for xf in styles.findall("cellXfs/xf"):
            numero = int(xf.get("numFmtId") or 0)
            code = personnalises.get(numero)
            if code is not None:
                formats.append(None if code.lower() == "general" else code)
            else:
                formats.append(numero or None)
        return formats


# --------------------------------------------------------------------------
# Ecriture d'un classeur .xlsx
# --------------------------------------------------------------------------
INTERDITS_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def xml_texte(texte: str) -> str:
    texte = INTERDITS_XML.sub("", texte)
    return texte.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def numero_date(jour: date) -> int:
    return (jour - date(1899, 12, 30)).days


STYLE_TITRE, STYLE_DATE = 1, 2


def ecrire_classeur(destination: Path, lignes_titres, lignes_donnees, largeurs) -> None:
    """Ecrit une feuille : lignes de titres (en gras, figees, filtrables) puis donnees.

    Une cellule est None ou (genre, valeur, format) ; genre "date" pour la colonne ajoutee.
    """
    chaines, formats = {}, {}

    def style_de(cle):
        if cle is None:
            return 0
        if cle not in formats:
            formats[cle] = 3 + len(formats)
        return formats[cle]

    def cellule_xml(reference, cellule, titre):
        genre, valeur, cle = cellule
        if genre == "s":
            indice = chaines.setdefault(valeur, len(chaines))
            attribut = ' s="1"' if titre else ""
            return f'<c r="{reference}" t="s"{attribut}><v>{indice}</v></c>'
        if genre == "date":
            return f'<c r="{reference}" s="{STYLE_DATE}"><v>{valeur}</v></c>'
        if genre == "b":
            vrai = 1 if valeur in ("1", "true") else 0
            return f'<c r="{reference}" t="b"><v>{vrai}</v></c>'
        style = STYLE_TITRE if titre else style_de(cle)
        attribut = f' s="{style}"' if style else ""
        return f'<c r="{reference}"{attribut}><v>{valeur}</v></c>'

    nb_colonnes = max([len(l) for l in lignes_titres] + [len(l) for l in lignes_donnees] + [1])
    rangees = []
    for numero, ligne in enumerate(list(lignes_titres) + list(lignes_donnees), 1):
        titre = numero <= len(lignes_titres)
        cellules = "".join(cellule_xml(f"{lettre_colonne(c)}{numero}", cellule, titre)
                           for c, cellule in enumerate(ligne) if cellule is not None)
        rangees.append(f'<row r="{numero}">{cellules}</row>')
    nb_lignes = len(rangees)
    if nb_lignes > LIGNES_MAX:
        raise Arret(f"{nb_lignes} lignes : c'est plus que ce qu'une feuille Excel peut contenir ({LIGNES_MAX}).")

    nb_titres = len(lignes_titres)
    derniere = lettre_colonne(nb_colonnes - 1)
    colonnes = "".join(f'<col min="{mini}" max="{maxi}" width="{largeur}" customWidth="1"/>'
                       for mini, maxi, largeur in largeurs)
    volet = (f'<pane ySplit="{nb_titres}" topLeftCell="A{nb_titres + 1}" activePane="bottomLeft" state="frozen"/>'
             if nb_titres else "")
    zone = f"A{nb_titres}:{derniere}{max(nb_lignes, nb_titres)}"
    filtre = f'<autoFilter ref="{zone}"/>' if nb_titres else ""
    feuille = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS}">'
        f'<dimension ref="A1:{derniere}{max(nb_lignes, 1)}"/>'
        f'<sheetViews><sheetView workbookViewId="0">{volet}</sheetView></sheetViews>'
        '<sheetFormatPr defaultRowHeight="15"/>'
        + (f"<cols>{colonnes}</cols>" if colonnes else "")
        + "<sheetData>" + "".join(rangees) + "</sheetData>" + filtre + "</worksheet>"
    )

    numeros, formats_perso = {}, ['<numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>']
    for cle in formats:  # dans l'ordre de creation des styles
        if isinstance(cle, int):
            numeros[cle] = cle
        else:
            numeros[cle] = 164 + len(formats_perso)
            code = xml_texte(cle).replace('"', "&quot;")
            formats_perso.append(f'<numFmt numFmtId="{numeros[cle]}" formatCode="{code}"/>')
    xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
           '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>',
           '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>']
    xfs += [f'<xf numFmtId="{numeros[cle]}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
            for cle in formats]
    styles = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS}">'
        + f'<numFmts count="{len(formats_perso)}">' + "".join(formats_perso) + "</numFmts>"
        + '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
          '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
          '<fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills>'
          '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        + f'<cellXfs count="{len(xfs)}">' + "".join(xfs) + "</cellXfs>"
        + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    )
    partage = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<sst xmlns="{NS}" count="{len(chaines)}" uniqueCount="{len(chaines)}">'
        + "".join(f'<si><t xml:space="preserve">{xml_texte(t)}</t></si>' for t in chaines)
        + "</sst>"
    )
    paquet = "http://schemas.openxmlformats.org/package/2006"
    office = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    types = "application/vnd.openxmlformats-officedocument.spreadsheetml"
    parties = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Types xmlns="{paquet}/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/xl/workbook.xml" ContentType="{types}.sheet.main+xml"/>'
            f'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="{types}.worksheet+xml"/>'
            f'<Override PartName="/xl/styles.xml" ContentType="{types}.styles+xml"/>'
            f'<Override PartName="/xl/sharedStrings.xml" ContentType="{types}.sharedStrings+xml"/></Types>',
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{paquet}/relationships">'
            f'<Relationship Id="rId1" Type="{office}/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<workbook xmlns="{NS}" xmlns:r="{office}">'
            '<sheets><sheet name="Mise au mille" sheetId="1" r:id="rId1"/></sheets>'
            + (('<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">'
                + "'Mise au mille'!" + re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", zone) + "</definedName></definedNames>")
               if nb_titres else "")
            + "</workbook>",
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="{paquet}/relationships">'
            f'<Relationship Id="rId1" Type="{office}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{office}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId3" Type="{office}/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": feuille,
        "xl/styles.xml": styles,
        "xl/sharedStrings.xml": partage,
    }
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as z:
            for nom, contenu in parties.items():
                z.writestr(nom, contenu.encode("utf-8"))
    except PermissionError:
        raise Arret(f"impossible d'ecrire {destination.name} : ferme ce fichier dans Excel puis relance.")


# --------------------------------------------------------------------------
# Titres de colonnes et alignement des fichiers entre eux
# --------------------------------------------------------------------------
def valeurs(ligne):
    return [c[1] if c else "" for c in ligne]


def remplies(ligne) -> int:
    return sum(1 for c in ligne if c)


def lignes_communes(a, b) -> int:
    """Nombre de lignes du haut identiques dans deux fichiers et ne contenant que du texte."""
    k = 0
    while (k < min(len(a), len(b), 6) and valeurs(a[k]) == valeurs(b[k])
           and all(c[0] == "s" for c in a[k] if c)):
        k += 1
    return k


def compter_lignes_de_titres(classeurs) -> int:
    """Nombre de lignes de titres en haut de chaque fichier (au moins 1).

    Ce sont les lignes identiques d'un jour au suivant. Un tableau a colonnes regroupees a une
    ligne de bandeaux au-dessus des titres ; a l'inverse, une ligne de regroupement des donnees
    ("Poste: Matin") peut se repeter chaque jour mais ne remplit qu'une case : la derniere ligne
    de titres doit donc etre une ligne bien remplie.
    """
    avec_donnees = [c.lignes for c in classeurs if c.lignes]
    comptes = [k for k in (lignes_communes(a, b) for a, b in zip(avec_donnees, avec_donnees[1:])) if k >= 1]
    if not comptes:
        return 1
    commun = max(set(comptes), key=lambda k: (comptes.count(k), -k))  # la valeur la plus frequente
    reference = avec_donnees[0]
    largeur = max(remplies(l) for l in reference[:commun])
    while commun > 1 and remplies(reference[commun - 1]) * 2 < largeur:
        commun -= 1
    return commun


def cles_de_colonnes(lignes_titres, largeur: int):
    """Identifie chaque colonne par ses titres (et son rang si deux colonnes ont le meme titre)."""
    cles, deja = [], {}
    for c in range(largeur):
        titres = tuple(l[c][1] if c < len(l) and l[c] else "" for l in lignes_titres)
        rang = deja.get(titres, 0)
        deja[titres] = rang + 1
        cles.append((titres, rang))
    return cles


# --------------------------------------------------------------------------
# Recherche des fichiers a assembler
# --------------------------------------------------------------------------
def dossier_telechargements() -> Path:
    if sys.platform == "win32":
        try:
            import winreg

            cle = r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, cle) as k:
                chemin = winreg.QueryValueEx(k, "{374DE290-123F-4565-9164-39C4925E467B}")[0]
            return Path(os.path.expandvars(chemin))
        except Exception:
            pass
    return Path.home() / "Downloads"


def trouver_source(chemins):
    """Renvoie la liste des sources : celles donnees, sinon le ZIP le plus recent, sinon le dossier du bot."""
    if chemins:
        return [Path(c.strip().strip('"')) for c in chemins]
    zips = []
    for dossier in (ICI, dossier_telechargements(), Path.home() / "Downloads", Path.home() / "Desktop"):
        try:
            zips += list(dossier.glob(MOTIF_ZIP))
        except OSError:
            pass
    if zips:
        return [max(zips, key=lambda f: f.stat().st_mtime)]
    if any((ICI / "extractions").glob("*.xls*")):
        return [ICI / "extractions"]
    print("Je ne trouve pas le fichier ZIP de l'extraction (Mise_au_mille_..._au_....zip).")
    print("Fais glisser ce fichier ZIP dans cette fenetre, puis appuie sur Entree :")
    try:
        saisie = input("> ").strip().strip('"')
    except EOFError:
        saisie = ""
    if not saisie:
        raise Arret("aucun fichier indique.")
    return [Path(saisie)]


def charger_fichiers(sources):
    """Renvoie [(date, nom, octets)] tries par date, a partir de ZIP, dossiers ou fichiers .xlsx."""
    trouves = {}

    def ajouter(nom: str, lire):
        m = MOTIF_DATE.search(Path(nom).name)
        if not m or not nom.lower().endswith((".xlsx", ".xlsm")) or Path(nom).name.startswith("~$"):
            return
        try:
            jour = date(int(m[1]), int(m[2]), int(m[3]))
        except ValueError:
            return
        trouves[jour] = (jour, Path(nom).name, lire())

    for source in sources:
        if not source.exists():
            raise Arret(f"introuvable : {source}")
        if source.is_dir():
            for f in sorted(source.glob("*.xls*")):
                ajouter(f.name, f.read_bytes)
        elif source.suffix.lower() == ".zip":
            try:
                with zipfile.ZipFile(source) as z:
                    for nom in z.namelist():
                        ajouter(nom, lambda nom=nom: z.read(nom))
            except zipfile.BadZipFile:
                raise Arret(f"{source.name} n'est pas un fichier ZIP valide (telechargement incomplet ?).")
        else:
            ajouter(source.name, source.read_bytes)
    return [trouves[jour] for jour in sorted(trouves)]


# --------------------------------------------------------------------------
# Assemblage
# --------------------------------------------------------------------------
def assembler(sources, sortie: Path, nb_titres_impose, avec_fichiers_du_jour: bool) -> Path:
    print("Source : " + ", ".join(str(s) for s in sources))
    fichiers = charger_fichiers(sources)
    if not fichiers:
        raise Arret("aucun fichier Excel avec une date dans son nom (AAAA-MM-JJ) n'a ete trouve.")
    print(f"{len(fichiers)} fichiers journaliers, du {fichiers[0][0]:%d/%m/%Y} au {fichiers[-1][0]:%d/%m/%Y}. Lecture...")

    jours, illisibles = [], []
    for jour, nom, octets in fichiers:
        try:
            jours.append((jour, nom, Classeur(octets)))
        except (Arret, KeyError, ValueError, IndexError, ET.ParseError) as e:
            illisibles.append(f"{nom} ({e})")
    if not jours:
        raise Arret("aucun fichier lisible : " + "; ".join(illisibles[:3]))

    nb_titres = nb_titres_impose or compter_lignes_de_titres([c for _, _, c in jours])
    print(f"Lignes de titres en haut de chaque fichier : {nb_titres}")

    # Colonnes du fichier unique : celles du premier fichier, puis celles apparues ensuite.
    colonnes, titres_union, apparues = {}, [], []
    plans = []
    for jour, nom, classeur in jours:
        titres = classeur.lignes[:nb_titres]
        largeur = max((len(l) for l in classeur.lignes), default=0)
        plan = []
        for c, cle in enumerate(cles_de_colonnes(titres, largeur)):
            if cle not in colonnes:
                colonnes[cle] = len(colonnes)
                titres_union.append([l[c] if c < len(l) else None for l in titres])
                if plans and any(cle[0]):
                    apparues.append(f"« {' / '.join(t for t in cle[0] if t)} » a partir du {jour:%d/%m/%Y}")
            plan.append(colonnes[cle])
        plans.append(plan)
    nb_colonnes = len(colonnes)

    def avec_date(jour, ligne, plan):
        sortie_ligne = [("date", numero_date(jour), None)] + [None] * nb_colonnes
        for c, cellule in enumerate(ligne):
            if cellule is not None:
                sortie_ligne[1 + plan[c]] = cellule
        return sortie_ligne

    lignes_titres = []
    for r in range(nb_titres):
        etiquette = ("s", TITRE_DATE, None) if r == nb_titres - 1 else None
        lignes_titres.append([etiquette] + [t[r] if r < len(t) else None for t in titres_union])

    donnees, lignes_par_jour = [], []
    for (jour, nom, classeur), plan in zip(jours, plans):
        corps = classeur.lignes[nb_titres:]
        lignes_par_jour.append((jour, len(corps)))
        donnees.extend(avec_date(jour, ligne, plan) for ligne in corps)

    reference = jours[0][2]
    largeurs = [(1, 1, "14")] + [(mini + 1, maxi + 1, largeur) for mini, maxi, largeur in reference.largeurs]
    complet = sortie / f"Mise_au_mille_complet_{jours[0][0]:%Y-%m-%d}_au_{jours[-1][0]:%Y-%m-%d}.xlsx"
    print("Ecriture du fichier unique...")
    ecrire_classeur(complet, lignes_titres, donnees, largeurs)

    if avec_fichiers_du_jour:
        print("Ecriture des fichiers journaliers avec la colonne date...")
        dossier = sortie / "fichiers_avec_date"
        for jour, nom, classeur in jours:
            titres = [[("s", TITRE_DATE, None) if r == nb_titres - 1 else None] + list(l)
                      for r, l in enumerate(classeur.lignes[:nb_titres])]
            corps = [[("date", numero_date(jour), None)] + list(l) for l in classeur.lignes[nb_titres:]]
            largeurs_jour = [(1, 1, "14")] + [(a + 1, b + 1, w) for a, b, w in classeur.largeurs]
            ecrire_classeur(dossier / (Path(nom).stem + ".xlsx"), titres, corps, largeurs_jour)

    print()
    print("=" * 62)
    print(f"Fichier unique : {complet}")
    print(f"  {len(donnees)} lignes de donnees, {nb_colonnes + 1} colonnes (dont « {TITRE_DATE} »), {len(jours)} jours")
    vides = [jour for jour, n in lignes_par_jour if n == 0]
    if vides:
        print(f"  Jours sans aucune ligne de donnees : {len(vides)} (" + ", ".join(f"{j:%d/%m/%Y}" for j in vides[:8])
              + (" ..." if len(vides) > 8 else "") + ")")
    if avec_fichiers_du_jour:
        print(f"Fichiers journaliers avec la colonne date : {sortie / 'fichiers_avec_date'}")
    if apparues:
        print("ATTENTION : les colonnes ne sont pas les memes tous les jours. Colonnes ajoutees en fin de tableau :")
        for texte in apparues[:15]:
            print("  - " + texte)
    if illisibles:
        print(f"ATTENTION : {len(illisibles)} fichiers illisibles, non repris : " + "; ".join(illisibles[:5]))
    return complet


def main() -> int:
    for flux in (sys.stdout, sys.stderr):
        try:
            flux.reconfigure(errors="replace")
        except (AttributeError, ValueError):
            pass
    parser = argparse.ArgumentParser(description="Assemble les Excel journaliers Mise au mille en un seul fichier.")
    parser.add_argument("sources", nargs="*", help="ZIP de l'extraction, dossier ou fichiers .xlsx (sinon : ZIP le plus recent)")
    parser.add_argument("--sortie", help="dossier de destination (par defaut : dossier 'assemblage' a cote de ce script)")
    parser.add_argument("--titres", type=int, help="nombre de lignes de titres en haut de chaque fichier (sinon detection)")
    parser.add_argument("--sans-fichiers-du-jour", action="store_true",
                        help="ne pas recreer chaque fichier journalier avec la colonne date")
    args = parser.parse_args()
    try:
        sortie = Path(args.sortie) if args.sortie else ICI / "assemblage"
        complet = assembler(trouver_source(args.sources), sortie, args.titres, not args.sans_fichiers_du_jour)
    except Arret as e:
        print(f"ARRET : {e}")
        return 2
    if sys.platform == "win32":
        try:
            os.startfile(str(complet.parent))  # ouvre le dossier du resultat dans l'Explorateur
        except OSError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
