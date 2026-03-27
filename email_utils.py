import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

def invia_email(destinatario, oggetto, corpo, allegato):
    mittente = 'tuo@email.com' # Sostituire con il proprio indirizzo
    password = 'password'      # Sostituire con la propria password
    msg = MIMEMultipart()
    msg['From'] = mittente
    msg['To'] = destinatario
    msg['Subject'] = oggetto
    msg.attach(MIMEText(corpo, 'plain'))
    with open(allegato, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename={allegato}')
        msg.attach(part)
    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(mittente, password)
    server.send_message(msg)
    server.quit()
