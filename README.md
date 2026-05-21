# vector-plus-daily-export

## Exécution quotidienne avec GitHub Actions

Ce dépôt contient un workflow GitHub Actions qui exécute chaque matin ces deux scripts :

- `scripts/export_vecteur_markets.py`
- `scripts/convert_json_to_table.py`

### Secrets GitHub requis (obligatoires)

- `VECTEUR_PLUS_PASSWORD` : Mot de passe Vecteur Plus

### Secrets GitHub optionnels (pour l'email)

- `SMTP_HOST` : Serveur SMTP (ex: `smtp.gmail.com`)
- `SMTP_PORT` : Port SMTP (ex: `587` pour Gmail)
- `SENDER_EMAIL` : Email expéditeur
- `SENDER_PASSWORD` : Mot de passe ou App Password
- `RECIPIENT_EMAIL` : Email destinataire
- `BASE_URL` : URL de base Vecteur Plus (optionnel)

**Si tu veux envoyer par email**, configure tous les secrets d'email.
**Si tu ne veux pas**, laisse-les vides (le workflow saute l'étape).

### Configuration M365 (Outlook)

1. Va dans GitHub → Settings → Secrets and variables → Actions
2. Ajoute ces secrets :
   - `SMTP_HOST` = `smtp.office365.com`
   - `SMTP_PORT` = `587`
   - `SENDER_EMAIL` = ton adresse M365 (ex: `user@company.com`)
   - `SENDER_PASSWORD` = ton mot de passe M365
   - `RECIPIENT_EMAIL` = l'email du destinataire (ex: `recipient@company.com`)

**Note** : Si ton organisation M365 a l'authentification multifacteur activée, tu dois utiliser un **App Password** (Passwords for app).

### Configuration Gmail (alternative)

1. Active l'authentification à deux facteurs sur ton compte Google
2. Génère un mot de passe d'application :
   - Va sur [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Sélectionne "Mail" et "Windows Computer"
   - Copie le mot de passe généré
3. Ajoute ces secrets :
   - `SMTP_HOST` = `smtp.gmail.com`
   - `SMTP_PORT` = `587`
   - `SENDER_EMAIL` = ton adresse Gmail
   - `SENDER_PASSWORD` = le mot de passe généré
   - `RECIPIENT_EMAIL` = l'email du destinataire

### Étapes à suivre

1. Pousse le projet dans un dépôt GitHub.
2. Crée les secrets GitHub (voir section Configuration Gmail ci-dessus).
3. Vérifie que le fichier `/.github/workflows/daily_export.yml` existe.

4. Le workflow est planifié pour s'exécuter tous les jours à `06:00 UTC`.
5. Il génère les fichiers suivants dans `responses/` :
   - `all_markets.json`
   - `all_markets.csv`
   - `all_markets.xlsx`
   - `all_markets_table.csv`
   - `all_markets_table.xlsx`
   - `all_markets_errors.json`

> Les fichiers CSV sont mis à jour en mode ajout si ils existent déjà. Le XLSX est reconstruit à partir du CSV cumulatif pour conserver les données précédentes.

6. Le workflow commite et pousse automatiquement ces fichiers dans la branche du dépôt si des changements sont détectés.
7. Les résultats sont également sauvegardés comme artefacts GitHub Actions.

### Ajuster l'horaire

- Change la ligne `cron: '0 6 * * *'` dans `/.github/workflows/daily_export.yml` si tu veux un autre horaire.
- Exemple pour 07:00 UTC : `cron: '0 7 * * *'`.
