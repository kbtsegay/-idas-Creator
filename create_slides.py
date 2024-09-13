import os
import base64
import datetime
import smtplib
from argparse import ArgumentParser
from email import encoders
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from google.oauth2 import service_account
from google.auth.transport.requests import Request
from google.auth.transport.urllib3 import AuthorizedHttp
from src.kidase_creator import KidaseCreator


def decode_credentials(encoded):
    # Decode the base64 string
    decoded = base64.b64decode(encoded)
    # Write the decoded content to credentials.json
    with open('credentials.json', 'wb') as file:
        file.write(decoded)


def send_email(file_path, recipient_email):
    try:
        # Load OAuth 2.0 credentials
        credentials = service_account.Credentials.from_service_account_file(
            'credentials.json',
            scopes=['https://www.googleapis.com/auth/gmail.send']
        )
        credentials.refresh(Request())
        
        # Create the email
        msg = MIMEMultipart()
        msg['From'] = 'your-email@gmail.com'
        msg['To'] = recipient_email
        msg['Subject'] = f'Your Kidase Slidedeck for {datetime.datetime.now().strftime("%B %d, %Y")}'

        part = MIMEBase('application', 'octet-stream')
        part.set_payload(open(file_path, 'rb').read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(file_path)}"')
        msg.attach(part)

        # Send the email
        authorized_http = AuthorizedHttp(credentials)
        response = authorized_http.request(
            'POST',
            'https://www.googleapis.com/gmail/v1/users/me/messages/send',
            body=msg.as_string()
        )
        print("Email sent successfully")
    except Exception as e:
        print(f"Error: {e}")


# Example usage
if __name__ == '__main__':
    parser = ArgumentParser(description='Decode the base64 encoded credentials')
    parser.add_argument('encoded', type=str, help='The base64 encoded credentials')
    args = parser.parse_args()

    decode_credentials(args.encoded)

    kidase_creator = KidaseCreator('./data', ['ግእዝ', 'ትግርኛ', 'english'])
    prs = kidase_creator.create_presentation()
    prs.save('test.pptx')
    
    send_email('test.pptx', 'kaleb.tsegay@gmail.com')
