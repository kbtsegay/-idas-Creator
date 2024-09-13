import os
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from src.kidase_creator import KidaseCreator

def send_email(file_path, recipient_email):
    try:
        email_address = os.getenv('EMAIL_ADDRESS')
        email_password = os.getenv('EMAIL_PASSWORD')
        
        # Debugging statements
        print(f"EMAIL_ADDRESS: {email_address}")
        print(f"EMAIL_PASSWORD: {email_password}")

        if email_address is None or email_password is None:
            raise ValueError("Email address or password environment variables are not set")

        msg = MIMEMultipart()
        msg['From'] = email_address
        msg['To'] = recipient_email
        msg['Subject'] = f'Your Kidase Slidedeck for {datetime.datetime.now().strftime("%B %d, %Y")}'

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(file_path, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)  # Update this line with the correct SMTP server
        server.starttls()
        server.login(email_address, email_password)
        server.sendmail(email_address, recipient_email, msg.as_string())
        server.quit()
        print("Email sent successfully")
    except Exception as e:
        print(f"Error: {e}")

# Example usage
if __name__ == '__main__':
    kidase_creator = KidaseCreator('./data', ['ግእዝ', 'ትግርኛ', 'english'])
    prs = kidase_creator.create_presentation()
    prs.save('test.pptx')
    send_email('test.pptx', 'kaleb.tsegay@gmail.com')

