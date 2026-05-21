import smtplib

smtp_host = "smtp.office365.com"
smtp_port = 587
sender_email = "gen-wkf-vorstone@groupesmac.onmicrosoft.com"
password = "V(538777937963ah"

with smtplib.SMTP(smtp_host, smtp_port, timeout=20) as server:
    server.set_debuglevel(1)
    server.ehlo()
    server.starttls()
    server.ehlo()
    server.login(sender_email, password)
    print("Connexion SMTP OK")