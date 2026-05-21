# vector-plus-daily-export

## Exécution quotidienne avec GitHub Actions

Ce dépôt contient un workflow GitHub Actions qui exécute chaque matin ces deux scripts :

- `scripts/export_vecteur_markets.py`
- `scripts/convert_json_to_table.py`

### Étapes à suivre

1. Pousse le projet dans un dépôt GitHub.
2. Crée un secret GitHub dans le dépôt :
   - `VECTEUR_PLUS_PASSWORD` : mot de passe de connexion Vecteur Plus
   - `VECTEUR_PLUS_BASE_URL` (facultatif) : URL de base de l'API si tu ne veux pas utiliser `app.http`
3. Vérifie que le fichier `/.github/workflows/daily_export.yml` existe.
4. Le workflow est planifié pour s'exécuter tous les jours à `06:00 UTC`.
5. Il génère les fichiers suivants dans `responses/` :
   - `all_markets.json`
   - `all_markets.csv`
   - `all_markets.xlsx`
   - `all_markets_table.csv`
   - `all_markets_table.xlsx`
   - `all_markets_errors.json`
6. Le workflow commite et pousse automatiquement ces fichiers dans la branche du dépôt si des changements sont détectés.
7. Les résultats sont également sauvegardés comme artefacts GitHub Actions.

### Ajuster l'horaire

- Change la ligne `cron: '0 6 * * *'` dans `/.github/workflows/daily_export.yml` si tu veux un autre horaire.
- Exemple pour 07:00 UTC : `cron: '0 7 * * *'`.
