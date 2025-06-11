import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


def send_email(smtp_server, smtp_port, username, password, sender, recipient, subject, body):
    """Send an email using the provided SMTP server credentials."""
    msg = MIMEMultipart()
    msg['From'] = sender
    msg['To'] = recipient
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))

    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        server.login(username, password)
        server.sendmail(sender, [recipient], msg.as_string())


if __name__ == "__main__":
    # Example usage
    SMTP_SERVER = "smtp.example.com"
    SMTP_PORT = 465
    USERNAME = "your_username"
    PASSWORD = "your_password"
    SENDER = "you@example.com"
    RECIPIENT = "recipient@example.com"
    SUBJECT = "Test Email"
    BODY = "This is a test email sent from the mail_sender script."

    send_email(SMTP_SERVER, SMTP_PORT, USERNAME, PASSWORD, SENDER, RECIPIENT, SUBJECT, BODY)
