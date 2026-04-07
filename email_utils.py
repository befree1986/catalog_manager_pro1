import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

def invia_email(destinatario, oggetto, corpo, allegato=None):
    # Inserisci qui il tuo indirizzo Gmail e la password di 16 caratteri generata
    mittente = 'tuo_account@gmail.com' 
    password = 'abcd efgh ilmn opqr'   
    
    msg = MIMEMultipart()
    msg['From'] = mittente
    msg['To'] = destinatario
    msg['Subject'] = oggetto
    msg.attach(MIMEText(corpo, 'plain'))
    
    if allegato and os.path.exists(allegato):
        with open(allegato, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(allegato)}')
            msg.attach(part)
            
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587, timeout=10)
        server.starttls()
        server.login(mittente, password)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        raise Exception(f"Errore di connessione o autenticazione: {str(e)}")
