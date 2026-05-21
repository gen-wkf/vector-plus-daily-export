#!/usr/bin/env python3
"""
Script pour envoyer le fichier XLSX généré via email.
"""

import sys
import smtplib
from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from datetime import datetime


def send_xlsx_email(
    smtp_host: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str,
    recipient_email: str,
    xlsx_file: str,
) -> None:
    """
    Envoie le fichier XLSX par email.

    Args:
        smtp_host: Serveur SMTP (ex: smtp.gmail.com)
        smtp_port: Port SMTP (ex: 587)
        sender_email: Email expéditeur
        sender_password: Mot de passe ou App Password
        recipient_email: Email destinataire
        xlsx_file: Chemin du fichier XLSX
    """
    xlsx_path = Path(xlsx_file)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Le fichier {xlsx_file} n'existe pas")

    # Créer le message
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = recipient_email
    message["Subject"] = f"Export Vecteur Plus - {datetime.now().strftime('%Y-%m-%d')}"

    # Corps du message
    body = f"""
Bonjour,

Veuillez trouver ci-joint l'export quotidien Vecteur Plus au format XLSX.

Date de génération: {datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')}
Fichier: {xlsx_path.name}

Cordialement,
Workflow GitHub Actions
"""
    message.attach(MIMEText(body, "plain", "utf-8"))

    # Ajouter le fichier XLSX en pièce jointe
    with open(xlsx_path, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {xlsx_path.name}")
    message.attach(part)

    # Envoyer l'email
    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(message)
        print(f"✓ Email envoyé avec succès à {recipient_email}")
    except smtplib.SMTPAuthenticationError:
        raise RuntimeError(
            f"Authentification échouée. Vérifiez le mot de passe et les paramètres SMTP."
        )
    except Exception as e:
        raise RuntimeError(f"Erreur lors de l'envoi : {e}")


if __name__ == "__main__":
    if len(sys.argv) < 7:
        print(
            "Usage: python send_export_mail.py <smtp_host> <smtp_port> "
            "<sender_email> <sender_password> <recipient_email> <xlsx_file>"
        )
        print("\nExemple (Gmail):")
        print(
            "  python send_export_mail.py smtp.gmail.com 587 "
            "your-email@gmail.com 'your-app-password' recipient@example.com export.xlsx"
        )
        sys.exit(1)

    smtp_host = sys.argv[1]
    smtp_port = int(sys.argv[2])
    sender_email = sys.argv[3]
    sender_password = sys.argv[4]
    recipient_email = sys.argv[5]
    xlsx_file = sys.argv[6]

    try:
        send_xlsx_email(
            smtp_host=smtp_host,
            smtp_port=smtp_port,
            sender_email=sender_email,
            sender_password=sender_password,
            recipient_email=recipient_email,
            xlsx_file=xlsx_file,
        )
    except Exception as e:
        print(f"Erreur: {e}", file=sys.stderr)
        sys.exit(1)
