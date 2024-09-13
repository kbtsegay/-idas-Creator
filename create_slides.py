import os
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from src.kidase_creator import KidaseCreator

def send_email(file_path, recipient_email):
    try:
        msg = MIMEMultipart()
        msg['From'] = os.getenv('EMAIL_ADDRESS')
        msg['To'] = recipient_email
        msg['Subject'] = f'Your Kidase Slidedeck for {datetime.datetime.now().strftime("%B %d, %Y")}'

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(file_path, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
        msg.attach(part)

        server = smtplib.SMTP('smtp.gmail.com', 587)  # Update this line with the correct SMTP server
        server.starttls()
        server.login(os.getenv('EMAIL_ADDRESS'), os.getenv('EMAIL_PASSWORD'))
        server.sendmail(os.getenv('EMAIL_ADDRESS'), recipient_email, msg.as_string())
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
