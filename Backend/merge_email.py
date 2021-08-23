##from main import Excel_file
import smtplib, ssl, email
import data
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_email(Receiver_Email, Attachment):
    '''
    Sends emails. Parameters are listed as such
    '''
    port = 465  # For SSL

    # Create a secure SSL context
    context = ssl.create_default_context()

    # Create a multipart message and set headers
    message = MIMEMultipart()
    message["From"] = data.Email
    message["To"] = Receiver_Email
    message["Subject"] = data.Subject
    message["Bcc"] = Receiver_Email

    # Add body to email
    message.attach(MIMEText(data.Message, "plain"))

    filename = str(Attachment)  # In same directory as script

    # Open PDF file in binary mode
    with open(filename, "rb") as attachment:
        # Add file as application/octet-stream
        # Email client can usually download this automatically as attachment
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )

    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()

    with smtplib.SMTP_SSL("smtp.gmail.com", port, context=context) as server:
        server.login(data.Email, data.Password)
        server.sendmail(data.Email, Receiver_Email, text)

