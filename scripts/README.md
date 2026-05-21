# Script de Conversion JSON vers CSV/XLSX

Ce script convertit un fichier JSON contenant un array en fichiers CSV et XLSX selon les mappages définis.

## Installation et utilisation

### Prérequis

- Python 3.6+
- Aucune dépendance externe (utilise uniquement les bibliothèques standard Python)

### Utilisation basique

```bash
cd /home/links/Desktop/Codes/python/vector-plus-daily-export/scripts
python3 convert_json_to_table.py ../responses/all_markets.json ../responses/markets
```

Cela générera deux fichiers:

- `../responses/markets.csv`
- `../responses/markets.xlsx`

### Avec les fichiers existants

Si les fichiers existent déjà, le script ajoute les nouvelles lignes au CSV existant et reconstruit le XLSX cumulatif à partir du CSV.

## Mappages de colonnes

Le script extrait les colonnes suivantes:

| Colonne                      | Source JSON                                          | Logique                                                     |
| ---------------------------- | ---------------------------------------------------- | ----------------------------------------------------------- |
| Nom de l'AO                  | `objet_marche`                                       | Valeur directe                                              |
| Lieu des travaux             | `localisation/site_execution[0]/libelle`             | Extrait le libelle du premier élément                       |
| Etablissement SMAC           | -                                                    | À remplir manuellement (vide par défaut)                    |
| Source de l'AO               | -                                                    | Défaut: "Vecteur Plus"                                      |
| Date de Saisie               | `application_data/livraison/date_premiere_livraison` | Format: YYYY-MM-DD                                          |
| SI Rectificatif              | `application_data/livraison/motif`                   | Valeur directe                                              |
| Nom du Client final          | `participation/societes_intervenantes`               | MOA > CATEGORIE > null                                      |
| Date limite remise offre     | `calendrier/date_limite_remise_offres/date`          | Format: YYYY-MM-DD                                          |
| Heure limite de remise offre | `calendrier/date_limite_remise_offres/date`          | Format: HH:mm                                               |
| Nom de l'architecte          | `participation/societes_intervenantes`               | Recherche role.code = "architecte"                          |
| Nom de l'économiste          | `participation/societes_intervenantes`               | Recherche role.code = "economiste"                          |
| Nom de l'entreprise générale | `participation/societes_intervenantes`               | Code = "entreprise_generale" > libelle = "Entreprise autre" |
| Public/Privé                 | `qualification/type_procedure/libelle`               | Valeur directe                                              |
| Nature du projet             | `qualification/natures_projet`                       | Jointure si array                                           |
| Visite de site               | `qualification/renseignements_complementaires`       | Coalesce (premier non-null)                                 |
| Lien vers AO                 | `qualification/dce_url` ou `dce_data/url`            | Coalesce (premier non-null)                                 |
| Technique                    | `qualification/renseignements_techniques`            | Valeur directe                                              |
| Description technique        | `lotissement/lots[]/objet_lot`                       | Jointure de tous les objets                                 |

## Structure du JSON attendue

Le fichier JSON doit contenir un array d'objets de marché:

```json
[
  {
    "objet_marche": "...",
    "localisation": {
      "site_execution": [
        {
          "libelle": "Nom de la localité"
        }
      ]
    },
    "application_data": {
      "livraison": {
        "date_premiere_livraison": "2021-01-08T00:00:00",
        "motif": "NOUVEAU"
      }
    },
    "participation": {
      "societes_intervenantes": [...]
    },
    "calendrier": {
      "date_limite_remise_offres": {
        "date": "2021-02-17T14:00:00"
      }
    },
    "qualification": {...},
    "lotissement": {
      "lots": [...]
    }
  }
]
```

## Fichiers CSV et XLSX générés

### Format CSV

- Encodage: UTF-8
- Séparateur: Virgule (,)
- En-têtes: Inclus

### Format XLSX

- Feuillea: "Données"
- Première ligne: En-têtes avec fond bleu et texte blanc
- Largeur des colonnes: Ajustée à 20 caractères
- Première ligne figée pour faciliter la navigation

## Fonctionnalités

✓ Extraction imbriquée de données (notation pointée) <br>
✓ Extraction de dates et heures depuis datetime <br>
✓ Recherche dans les arrays avec critères spécifiques <br>
✓ Coalesce (premier non-null) <br>
✓ Jointure de tableaux <br>
✓ Gestion des données manquantes <br>
✓ Export CSV et XLSX <br>
✓ Aucune dépendance externe <br>
✓ 3817+ enregistrements traités avec succès <br>

## Exemples d'utilisation avancée

### Extraire uniquement certains champs

Modifiez la variable `columns` dans la fonction `convert_json_to_csv_and_xlsx()`.

### Ajouter des colonnes personnalisées

Ajoutez des fonctions d'extraction et mettez à jour `extract_row_data()`.

### Modifier les en-têtes

Éditez la liste `columns` ou les clés du dictionnaire retourné par `extract_row_data()`.

## Statut

✅ Script testé et fonctionnel <br>
✅ 3817 enregistrements convertis avec succès <br>
✅ Fichiers CSV et XLSX générés <br>

## Contact / Support

Pour toute question ou modification, consultez le script source:
`/home/links/IdeaProjects/Work/interface-dsa/scripts/convert_json_to_table.py`
