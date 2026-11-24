# Bot d'extraction « Mise au mille »

Extrait automatiquement un fichier Excel par jour depuis la page « Mise au mille »
et le renomme avec la date du programme : `Mise_au_mille_2025-07-01.xlsx`, etc.

Le bot ouvre Edge (déjà présent sur Windows), règle la date, clique sur le bouton
d'extraction, range le fichier dans le dossier `extractions`, puis passe au jour suivant.

Il est adapté aux composants DevExtreme de la page (sélecteur de date `dxDateBox`, bouton
d'export du tableau) : la date est saisie au clavier ou, si le champ ne l'accepte pas,
cliquée dans le calendrier, puis contrôlée sur la vraie valeur du sélecteur avant chaque
extraction.

## Mise en route (Windows)

1. **Une seule fois** : double-clic sur `1_installer.bat` (installe le module Python `playwright`).
2. **Essai** : double-clic sur `2_test.bat`. Le bot traite seulement les 5 derniers jours.
   **Tout se passe dans la fenêtre Edge qui s'ouvre**, pas dans la fenêtre noire :
   - Connecte-toi si l'appli le demande. Le bot attend tout seul que la page « Mise au mille »
     soit affichée.
   - Il entoure en rouge le champ date et le bouton d'extraction, et un bandeau rouge en bas
     de la page demande « Est-ce correct ? ». Clique sur **Oui, lancer**.
   - S'il se trompe, clique sur **Non, je vais les montrer** : le bandeau te demande alors de
     cliquer toi-même sur le champ date, puis sur le bouton. Ces clics servent seulement à lui
     montrer les éléments, rien n'est déclenché.
   - Pendant le traitement, une petite étiquette verte en bas à droite indique l'avancement.
     Quand c'est fini, Edge se ferme et le bilan s'affiche dans la fenêtre noire.
   - Ouvre deux ou trois fichiers du dossier `extractions` pour vérifier que chaque fichier
     contient bien le programme du jour indiqué dans son nom.
3. **Extraction complète** : double-clic sur `3_lancer_extraction.bat`
   (du 1er juillet 2025 à aujourd'hui). Ne ferme pas la fenêtre Edge pendant le traitement.

Les bandeaux jaunes d'Edge du type « contrôlé par un logiciel de test automatisé » sont normaux.

## À savoir

- **Reprise** : les jours déjà extraits sont sautés. En cas de coupure, relance simplement
  `3_lancer_extraction.bat`.
- **Jours sans production** : si le tableau du jour est vide, le bot n'extrait rien (le fichier
  ne contiendrait que les titres de colonnes), note le jour « aucun fichier » et continue.
  La liste de ces jours est affichée à la fin. Pour extraire quand même ces fichiers vides,
  mets `"exporter_si_vide": true` dans `config.json`.
- **Rapport** : `rapport_extraction.csv` (s'ouvre dans Excel) liste chaque jour traité.
- **Sécurité** : si les premiers fichiers extraits ont tous le même contenu, le bot s'arrête
  et les met de côté dans `extractions\_suspects` : c'est le signe que la page n'a pas pris en
  compte la date saisie.

## Options (ligne de commande)

```
python extraction.py --debut 2025-07-01 --fin 2025-07-31   période précise
python extraction.py --test                                 5 derniers jours
python extraction.py --reapprendre                          remontrer le champ date et le bouton
python extraction.py --diagnostic                           décrire la page sans rien extraire
```

## Réglages (`config.json`)

| Clé | Rôle |
|---|---|
| `url` | adresse de la page « Mise au mille » |
| `date_debut` | premier jour à extraire |
| `dossier_sortie` | dossier des fichiers Excel |
| `nom_fichier`, `format_date_fichier` | nom des fichiers, par ex. `"%d-%m-%Y"` pour `Mise_au_mille_01-07-2025` |
| `navigateur` | `msedge`, `chrome` ou `chromium` |
| `selecteur_valider` | bouton à cliquer après avoir changé la date, s'il y en a un |
| `exporter_si_vide` | `true` pour extraire aussi les jours où le tableau est vide |
| `pause_apres_date` | pause en secondes avant de cliquer sur l'extraction (à augmenter si la page est lente) |
| `delai_sans_reponse` | secondes sans téléchargement avant de conclure « aucun fichier » |

## En cas de problème

- **Dans tous les cas** : le fichier `journal.log` du dossier contient tout ce que le bot a fait.
  C'est le fichier à envoyer pour comprendre un blocage.
- **Edge s'ouvre mais il ne se passe rien, puis il se referme** : le PC interdit de piloter Edge
  (politique de l'entreprise). Le bot l'indique dans la fenêtre noire. Lance alors
  `4_installer_chromium.bat` : il télécharge un navigateur à part, réservé au bot, puis relance
  `2_test.bat`.
- **Le bot n'arrive pas à régler la date ou s'arrête** : il crée un dossier `diagnostic`.
  Le fichier `diagnostic\elements.txt` (liste des champs et boutons de la page) suffit en
  général pour l'ajuster.
- **`pip` ne peut rien télécharger** (proxy d'entreprise) : il faut installer le module
  `playwright` par le moyen habituel de l'entreprise.
