# Bot d'extraction « Mise au mille »

Extrait automatiquement un fichier Excel par jour depuis la page « Mise au mille »
et le renomme avec la date du programme : `Mise_au_mille_2025-07-01.xlsx`, etc.

Le bot ouvre Edge (déjà présent sur Windows), règle la date, clique sur le bouton
d'extraction, range le fichier dans le dossier `extractions`, puis passe au jour suivant.

## Mise en route (Windows)

1. **Une seule fois** : double-clic sur `1_installer.bat` (installe le module Python `playwright`).
2. **Essai** : double-clic sur `2_test.bat`. Le bot traite seulement les 5 derniers jours.
   - Une fenêtre Edge s'ouvre. Connecte-toi si l'appli le demande, puis appuie sur Entrée
     dans la fenêtre noire.
   - Le bot entoure en rouge le champ date et le bouton d'extraction qu'il a trouvés et
     demande confirmation. S'il se trompe (ou ne trouve rien), réponds `n` : un bandeau rouge
     en bas de la page te demande alors de cliquer toi-même sur le champ date, puis sur le
     bouton. Ces clics servent seulement à lui montrer les éléments, rien n'est déclenché.
   - Ouvre deux ou trois fichiers du dossier `extractions` pour vérifier que chaque fichier
     contient bien le programme du jour indiqué dans son nom.
3. **Extraction complète** : double-clic sur `3_lancer_extraction.bat`
   (du 1er juillet 2025 à aujourd'hui). Laisse la fenêtre Edge ouverte pendant le traitement.

## À savoir

- **Reprise** : les jours déjà extraits sont sautés. En cas de coupure, relance simplement
  `3_lancer_extraction.bat`.
- **Jours sans production** : si aucun fichier n'est proposé, le jour est noté « aucun fichier »
  et le bot continue. La liste est affichée à la fin.
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
| `pause_apres_date` | pause en secondes avant de cliquer sur l'extraction (à augmenter si la page est lente) |
| `delai_sans_reponse` | secondes sans téléchargement avant de conclure « aucun fichier » |

## En cas de problème

- **Le bot n'arrive pas à régler la date ou s'arrête** : il crée un dossier `diagnostic`.
  Le fichier `diagnostic\elements.txt` (liste des champs et boutons de la page) suffit en
  général pour l'ajuster.
- **`pip` ne peut rien télécharger** (proxy d'entreprise) : il faut installer le module
  `playwright` par le moyen habituel de l'entreprise.
- **Edge refuse de s'ouvrir** (pilotage bloqué par la politique de l'entreprise) :
  `python -m playwright install chromium`, puis mettre `"navigateur": "chromium"` dans `config.json`.
